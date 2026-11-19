# Phase 1：基础设施

**依赖：Phase 0 已完成。**

## 1. 阶段目的

建立所有后续模块共用的基础能力：本地日志、统一错误表达、多工作簿配置读写、UNC 路径校验与只读诊断、共享路径延迟探测，以及可供 `MdlDiff`/`MdlMerge` 共用的单元格值、字体格式和填充比较。完成后，配置支持登记多份 UNC 工作簿，同时兼容读取旧版单路径配置。

## 2. 实现方式

- 在 `MdlInfra` 中定义稳定的错误码、日志级别和轻量错误信息结构。公共过程返回成功标志与错误信息，不让底层错误静默消失。
- 配置读写支持一组 UNC 工作簿路径；每条路径单独规范化，遇到盘符路径、空值、无效 UNC 或不区分大小写后的重复路径时拒绝配置并说明原因。不提供盘符到 UNC 的自动解析。
- 配置保存在 `%LOCALAPPDATA%\MDDTLock\config.ini`，由用户手动创建和维护，没有配置窗体。首次使用时通过文件资源管理器定位 `%LOCALAPPDATA%\MDDTLock`，必要时创建目录；用纯文本编辑器创建/编辑 `config.ini`。每个工作簿一条 `WorkbookPath=`，新增时加行，移除时删行，保留至少一条有效 UNC 路径。版本 2 使用重复的 `WorkbookPath=` 登记多路径；读取器兼容版本 1 的单条 `MainFilePath=`，写入器统一写版本 2。缺少配置时不猜测、不自动填入共享路径。
- 提供只读 UNC 路径诊断，将可判定的文件/目录不存在、访问拒绝和网络路径不可达区分为稳定状态，并保留底层系统错误码；诊断不创建共享文件。无法可靠区分的状态保留为“无法确认”，不得误报为文件不存在。
- 延迟探测只对已有共享文件或目录做只读访问计时，不在共享上创建探测文件。探测失败要返回错误状态，不把失败解释成“共享文件不存在”。
- 业务值比较规则集中在 `MdlInfra`，避免 `MdlDiff` 与 `MdlMerge` 各自实现后出现语义分叉：非空值必须为字符串，字符串逐字符精确比较；`Empty` 作为空白哨兵，仅与 `Empty` 相等，零长度字符串 `""` 与 `Empty` 不同。其它业务值 Variant 类型返回输入错误。
- 字体与填充快照在 Excel 边界被编码为版本化、长度前缀的纯字符串；`FP1` 的 14 个字段依次为 `Name`、`Size`、`Bold`、`Italic`、`Underline`、`Strikethrough`、`Superscript`、`Subscript`、`OutlineFont`、`Shadow`、`ColorIndex`、`Color`、`ThemeColor`、`TintAndShade`；`FL1` 的 9 个字段依次为 `Pattern`、`ColorIndex`、`Color`、`ThemeColor`、`TintAndShade`、`PatternColorIndex`、`PatternColor`、`PatternThemeColor`、`PatternTintAndShade`。`TF1` 绑定 `Empty`/字符串标记、原文和压缩字体 runs；每个 run 为 `[一基起点, 长度, FP1 签名]`，连续相同签名合并且 runs 必须覆盖完整字符串。空白/零长度文本无字符 run，保存单元格级字体签名。编码字段保留类型、长度和原值，不允许依赖分隔符转义或区域格式解析。
- `CompareFormattingData` 只接受可完整解析、组件类型相同的规范化字符串，并逐字节精确比较；它不递归比较任意数组。公开解析/构造过程用于在 Excel 边界生成、校验及展示纯数据，不接收 Excel 对象。格式值只允许字符串、数值、布尔值、Empty 和内部约定的“不适用”错误哨兵；Null、对象、数组或未知类型拒绝。字符文本与其字体映射由上层作为一个 `TextFont` 组件比较，填充作为另一组件比较。
- 日志写入用户本机 `%LOCALAPPDATA%` 下的应用目录，避免业务数据和运行文件写入共享；具体日志保留策略在风险项中待确认。
- `MdlInfra` 不依赖其它产品模块，不引用 Excel 对象模型。

## 3. 涉及函数与模块

为后续会话目录和 fencing 标记提供本地路径守卫与 GUID 生成：EnsureLocalAppDirectory 只允许在 LOCALAPPDATA 下建立目录；IsLocalAppDataPath 只做路径归属校验；NewGuidString 调用 Windows GUID API，生成失败返回空字符串，由调用方立即中止，不退化成低熵随机值。

模块：`src/MdlInfra/`，不新增产品层依赖。

当前过程接口：

```text
Public Sub LogWrite(ByVal level As LogLevel, ByVal source As String, ByVal message As String)
Public Function MakeAppError(ByVal code As Long, ByVal source As String, ByVal message As String, Optional ByVal context As String = "") As AppErrorInfo
Public Function NewGuidString() As String
Public Function EnsureLocalAppDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function IsLocalAppDataPath(ByVal path As String) As Boolean
Public Function GetLocalAppDataFolder() As String
Public Function GetDefaultConfigPath() As String
Public Function ValidateUncPath(ByVal path As String, ByRef normalizedPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ReadConfig(ByVal configPath As String, ByRef config As AppConfig, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function WriteConfig(ByVal configPath As String, ByRef config As AppConfig, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function InspectUncPath(ByVal path As String, ByRef status As SharedPathStatus, ByRef systemError As Long, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ClassifySharedPathError(ByVal systemError As Long) As SharedPathStatus
Public Function ProbeUncLatency(ByVal path As String, ByRef elapsedMs As Double, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function CompareCellValues(ByVal leftValue As Variant, ByVal rightValue As Variant, ByRef equal As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function EncodeFormatField(ByVal value As String) As String
Public Function DecodeFormatField(ByRef data As String, ByRef cursor As Long, ByRef value As String) As Boolean
Public Function BuildFontSignature(ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As String
Public Function BuildTextFontData(ByVal cellValue As Variant, ByRef runs As Variant, ByVal emptyCellFont As String, ByRef errorInfo As AppErrorInfo) As String
Public Function BuildFillData(ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As String
Public Function ParseTextFontData(ByVal data As String, ByRef valueKind As String, ByRef textValue As String, ByRef runs As Variant, ByRef emptyCellFont As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ParseFontSignature(ByVal data As String, ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ParseFillData(ByVal data As String, ByRef properties As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function DecodeFormatScalar(ByVal encodedValue As String, ByRef value As Variant, ByRef isApplicable As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function CompareFormattingData(ByVal leftValue As Variant, ByVal rightValue As Variant, ByRef equal As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function FormatFormattingPreview(ByVal data As Variant) As String
Private Function NormalizeConfigPaths(ByRef sourcePaths As Variant, ByRef normalizedPaths As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function EnsureLocalDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function EnsureDirectoryTree(ByVal path As String) As Boolean
Private Function IsUnderLocalAppData(ByVal path As String) As Boolean
```

`AppConfig.WorkbookPaths` 保存有序 UNC 路径数组，`Version` 保存配置版本，不保存显示偏好或业务状态。`config.ini` 的版本 2 示例：

```ini
Version=2
WorkbookPath=\\server\share\folder-a\workbook-a.xlsx
WorkbookPath=\\server\share\folder-b\workbook-b.xlsx
```

上述路径是格式示例，不是可用配置。用户必须替换为实际 UNC 路径。映射盘符如 `X:\folder\book.xlsx` 不接受，因为不同客户端可能将其映射到不同位置，导致同一共享文件的锁标识分裂。`AppErrorInfo` 至少表达模块来源、稳定错误码和面向上层的说明。

`SharedPathStatus` 区分 `SharedPathAvailable`、`SharedPathNotFound`、`SharedPathAccessDenied`、`SharedPathNetworkUnavailable`、`SharedPathIsDirectory` 和 `SharedPathUnknown`；底层 API 无法可靠判别时返回 `Unknown` 并保留系统错误码。Windows 文件属性查询只读，不会创建共享文件。

## 4. 功能范围

做：本地日志、错误封装、多 UNC 路径校验与只读诊断、版本 1/2 配置读写、只读延迟探测、统一字符串/空值比较、规范化字体/填充快照构造/解析/比较/摘要预览、纯逻辑单元测试。

本阶段不做 Ribbon/设置窗体、会话状态、共享锁、工作簿读写、共享审计日志、自动寻找或改写业务文件。不得接受映射盘符作为配置，不得为探测在共享上创建临时文件。

## 5. 测试方案

- 覆盖 LOCALAPPDATA 路径守卫对父目录穿越的拒绝，并检查 Windows GUID 格式与连续生成不重复。配置拒绝路径测试不得创建目标文件。

- `test/MdlInfra/`：UNC 有效/无效形式、盘符拒绝写入配置、多个路径读写、版本 1 兼容、重复路径、配置缺失/损坏、路径诊断状态映射、日志目录不可写、字符串精确比较、大小写/空格差异、`Empty` 与 `""`、长度前缀含分隔符、空文本字体状态、run 连续覆盖、固定字段顺序、填充/字体类型不匹配、损坏记录及格式摘要预览。VBA 测试由 `test/TestSupport.bas` 捕获断言错误并返回给 `tools/Run-VbaTests.ps1`，避免隐藏 Excel 因未处理错误停在模态对话框；测试逻辑不访问工作簿对象。
- 路径和配置测试使用内存或本机临时目录；延迟探测的共享路径测试单独标记环境前提，且只测已有目标。
- 覆盖空字符串、尾部分隔符、大小写、权限不足和 I/O 失败；验证日志失败不会覆盖原始业务错误。
- 测试输入只使用内存数据或 `%LOCALAPPDATA%` 下的本机临时配置；不打开 Excel 工作簿、不访问 Excel 对象模型，也不访问真实共享。当前 VBA runner 使用受控隐藏 Excel 作为 VBA 执行宿主，并设置硬超时及进程清理。

## 6. 风险与未决问题

- **日志保存位置与保留周期未由总计划确定。** 选项：本机 `%LOCALAPPDATA%` 滚动日志，或只保留短期诊断日志。暂按本机应用目录、限制体积和不记录业务单元格内容处理；需确认保留周期及是否需要用户可见的清理入口。
- **配置由用户手工维护。** 当前加载项没有配置管理窗体；用户需在 `%LOCALAPPDATA%\MDDTLock\config.ini` 编辑版本 2 清单。错误或缺失时显示文件位置和格式要求，不自动猜测、创建或修复共享路径。
- **Windows 对共享路径错误的分类可能因服务器、权限和 SMB 状态而不同。** 暂定只对明确的系统错误分类；其余保留原始系统错误并显示“无法确认”，需在目标 Samba 环境验收。
- **字符串比较暂定区分大小写且保留空格，`Empty` 与 `""` 区分。** 这是精确比较的保守解释；若业务把大小写或空白归一视为相等，需回来确认后调整共享比较函数及 Diff/Merge 测试。
- 纯字符串协议已由内存测试验证；其 Excel 属性读取、主题/自动颜色、填充图案及写入后等价状态仍依赖目标 Excel 的真实对象模型。Phase 6/9 必须覆盖这些映射，无法无损表示时应失败关闭，不得静默丢弃有效格式。
- 延迟探测只能说明某次文件系统访问耗时，不能保证随后的提交速度；界面不得把它当作可用性承诺。
