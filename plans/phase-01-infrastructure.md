# Phase 1：基础设施

**依赖：Phase 0 已完成。**

## 1. 阶段目的

建立所有后续模块共用的基础能力：本地日志、统一错误表达、配置读写、UNC 路径校验与路径解析、共享路径延迟探测，以及可供 `MdlDiff`/`MdlMerge` 共用的字符串单元格值比较。完成后，调用方能够得到可记录、可展示、可测试的明确结果；配置始终保存 UNC 路径，不依赖映射盘符。

## 2. 实现方式

- 在 `MdlInfra` 中定义稳定的错误码、日志级别和轻量错误信息结构。公共过程返回成功标志与错误信息，不让底层错误静默消失。
- 配置读写支持一组 UNC 工作簿路径；每条路径单独规范化，遇到盘符路径、空值或重复路径时拒绝配置并说明原因。另提供路径解析能力供诊断或迁移输入使用，但解析结果必须经过 UNC 校验后才能成为配置值。
- 配置保存在 `%LOCALAPPDATA%\MDDTLock\config.ini`。当前版本需手工创建和维护；多路径扩展暂按手工维护处理。版本 2 使用重复的 `WorkbookPath=` 项登记多个 UNC 路径。读取器兼容版本 1 的单条 `MainFilePath=`，写入器统一写版本 2。缺少配置时不猜测、不自动填入共享路径。是否由 Phase 8 增加配置窗体仍待确认。
- 提供只读 UNC 路径诊断，将可判定的文件/目录不存在、访问拒绝和网络路径不可达区分为稳定状态，并保留底层系统错误码；诊断不创建共享文件。无法可靠区分的状态保留为“无法确认”，不得误报为文件不存在。
- 延迟探测只对已有共享文件或目录做只读访问计时，不在共享上创建探测文件。探测失败要返回错误状态，不把失败解释成“共享文件不存在”。
- 单元格比较规则集中在 `MdlInfra`，避免 `MdlDiff` 与 `MdlMerge` 各自实现后出现语义分叉：非空值必须为字符串，字符串逐字符精确比较；`Empty` 作为空白哨兵，仅与 `Empty` 相等，零长度字符串 `""` 与 `Empty` 不同。其它 Variant 类型返回输入错误。
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
Public Function ResolvePathToUnc(ByVal path As String, ByRef uncPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ReadConfig(ByVal configPath As String, ByRef config As AppConfig, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function WriteConfig(ByVal configPath As String, ByRef config As AppConfig, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function InspectUncPath(ByVal path As String, ByRef status As SharedPathStatus, ByRef systemError As Long, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ProbeUncLatency(ByVal path As String, ByRef elapsedMs As Double, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function CompareCellValues(ByVal leftValue As Variant, ByVal rightValue As Variant, ByRef equal As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function EnsureLocalDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function EnsureDirectoryTree(ByVal path As String) As Boolean
Private Function IsUnderLocalAppData(ByVal path As String) As Boolean
```

`AppConfig` 保存有序 UNC 路径数组及配置版本，不保存显示偏好或业务状态。`config.ini` 的版本 2 示例：

```ini
Version=2
WorkbookPath=\\server\share\folder-a\workbook-a.xlsx
WorkbookPath=\\server\share\folder-b\workbook-b.xlsx
```

上述路径是格式示例，不是可用配置。用户必须替换为实际 UNC 路径。`AppErrorInfo` 至少表达模块来源、稳定错误码和面向上层的说明。

`SharedPathStatus` 至少区分 `Available`、`NotFound`、`AccessDenied`、`NetworkUnavailable` 和 `Unknown`；底层 API 无法可靠判别时返回 `Unknown` 并保留系统错误码。

## 4. 功能范围

做：本地日志、错误封装、多 UNC 路径校验/解析与只读诊断、版本 1/2 配置读写、只读延迟探测、统一字符串/空值比较、纯逻辑单元测试。

本阶段不做 Ribbon/设置窗体（界面范围由 Phase 8 计划决定）、会话状态、共享锁、工作簿读写、共享审计日志、自动寻找或改写业务文件。不得写入映射盘符作为配置，不得为探测在共享上创建临时文件。

## 5. 测试方案

- 覆盖 LOCALAPPDATA 路径守卫对父目录穿越的拒绝，并检查 Windows GUID 格式与连续生成不重复。配置拒绝路径测试不得创建目标文件。

- `test/MdlInfra/`：UNC 有效/无效形式、盘符拒绝写入配置、多个路径读写、版本 1 兼容、重复路径、UNC 解析失败、配置缺失/损坏、路径诊断状态映射、日志目录不可写、字符串精确比较、大小写/空格差异、`Empty` 与 `""`、数字/布尔/错误值拒绝。VBA 测试由 `test/TestSupport.bas` 捕获断言错误并返回给 `tools/Run-VbaTests.ps1`，避免隐藏 Excel 因未处理错误停在模态对话框；测试逻辑不访问工作簿对象。
- 路径和配置测试使用内存或本机临时目录；延迟探测的共享路径测试单独标记环境前提，且只测已有目标。
- 覆盖空字符串、尾部分隔符、大小写、权限不足和 I/O 失败；验证日志失败不会覆盖原始业务错误。
- 测试输入只使用内存数据或 `%LOCALAPPDATA%` 下的本机临时配置；不打开 Excel 工作簿、不访问 Excel 对象模型，也不访问真实共享。当前 VBA runner 使用受控隐藏 Excel 作为 VBA 执行宿主，并设置硬超时及进程清理。

## 6. 风险与未决问题

- **日志保存位置与保留周期未由总计划确定。** 选项：本机 `%LOCALAPPDATA%` 滚动日志，或只保留短期诊断日志。暂按本机应用目录、限制体积和不记录业务单元格内容处理；需确认保留周期及是否需要用户可见的清理入口。
- **多路径配置的管理方式尚未最终确认。** 当前加载项没有设置窗体，扩展阶段暂按用户手工维护处理，并提供格式示例和配置缺失/损坏提示；需确认是否接受手工编辑，或应增加图形化配置入口。无论采用哪种入口，都不自动猜测、创建或修复共享路径。
- **Windows 对共享路径错误的分类可能因服务器、权限和 SMB 状态而不同。** 暂定只对明确的系统错误分类；其余保留原始系统错误并显示“无法确认”，需在目标 Samba 环境验收。
- **盘符到 UNC 的解析依赖 Windows 网络映射状态。** 暂定解析失败即报错，不猜测服务器路径；配置入口仍只接受 UNC。需在目标 Windows 环境验证解析 API 的可用性和权限行为。
- **字符串比较暂定区分大小写且保留空格，`Empty` 与 `""` 区分。** 这是精确比较的保守解释；若业务把大小写或空白归一视为相等，需回来确认后调整共享比较函数及 Diff/Merge 测试。
- 延迟探测只能说明某次文件系统访问耗时，不能保证随后的提交速度；界面不得把它当作可用性承诺。
