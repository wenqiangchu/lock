# Phase 6：Excel 适配层

**依赖：Phase 2、Phase 3。** 与 Phase 5 可并行，双方通过 `MdlStorage` 接口契约衔接。

## 1. 阶段目的

建立唯一接触 Excel 对象模型的适配层。完成后，上层只使用会话句柄、schema 和内存数组即可打开/读取/写入/保存工作簿，不会持有或传递 `Workbook`、`Worksheet`、`Range` 对象。

## 2. 实现方式

- `MdlExcel` 使用宿主 Excel 实例，管理本地副本和当前会话所选共享工作簿的打开句柄；句柄以内部 ID 暴露给上层，Excel 对象引用只保存在模块私有状态中。每次会话只打开一份共享工作簿。
- 打开配置中的共享文件前调用 `MdlInfra.InspectUncPath` 做只读预检。明确的文件/目录不存在、访问拒绝、网络路径不可达及配置路径实际是文件夹分别返回稳定错误码；`Workbooks.Open` 失败时保留 Excel 错误号和描述，不从通用 COM 错误推断锁状态。成功打开但 `Workbook.ReadOnly=True` 时作为“无法取得写入访问”单独报告，并提示可能是其他用户持有 Excel 锁、文件只读属性或当前账号权限不足；当前实现不单独识别 sharing violation。
- 每次快照调用只处理一个 sheet。目标参数是 N×2 Variant 矩阵（sheet 名、AMG 表头）；快照数据矩阵按物理行号和物理列号索引，包含所有物理列至最后一个非空表头。内部空行保留原行号。
- `SnapshotData` 保存 sheet 名、完整表头指纹、数据首末行、目标、声明列索引、A–E 锚点列索引、锚点值矩阵、单元格值矩阵、逐字符直接字体格式矩阵和单元格直接填充矩阵。值矩阵覆盖物理列，字体/填充矩阵仅读取声明列。字体矩阵按相邻属性相同的字符段压缩为 runs；每个 run 保存一基字符起点、长度和 `FP1` 字段顺序固定的直接 Font 数据，所有字符恰好覆盖一次。空值或零长度字符串没有字符 run 时保存单元格级 Font 状态。填充矩阵保存每个声明单元格字段顺序固定的直接 Interior 状态；`FP1`/`FL1` 字段顺序以 Phase 1 定义为准。
- 共享主文件复制到 `%LOCALAPPDATA%` 工作区，由 `MdlStorage` 执行文件操作。Base 快照从本地副本在签出时读取；提交时重新读取共享文件形成 Remote。
- 读取表头后调用 `MdlSchema` 解析。用有界读取策略取得声明列、product 锚点值及目标行数组；保留物理行号，以 product 锚点列值验证 Base、Local、Remote 行映射，不压缩内部全空行。禁止 `Range.Find` 和 `UsedRange` 求边界。候选末行方法需在代表性模板或受控 Excel 样例中验证并记录。
- 签出和提交前验证声明 AMG 列中没有公式，非空单元格均为字符串；发现公式或其他数据类型即中止并报告单元格位置。该检查仅针对本次声明列。格式读取只接受可确定读回的直接字体和填充状态；无法表示的混合/自动状态必须明确拒绝，不能静默丢失。`ThemeColor=0` 时把主题色及对应色调归一为“不适用”；`Pattern` 为 `xlPatternSolid` 或 `xlPatternNone` 时把图案色及其主题/色调字段归一为“不适用”，避免比较不可见字段被 Excel 自动重写。
- 提交前再次校验 sheet、表头/列名指纹、锚点列及其顺序/值映射、数组尺寸。若锚点被编辑、行增删/移动或结构不符，失败关闭，不进行写入。
- 写回只针对会话声明的 AMG 列，分别写入合并后的值、字符字体 runs 和单元格填充；按声明列写值，不以 `Union` 操作非连续区域，也不把整张工作表的格式复制回去。值批量写入后逐格读回校验；私有 `WriteTextFontData` 和 `WriteFillData` 各自在设置后重新采集规范化快照并与请求数据比较，任何属性无法往返均失败关闭。校验通过后对共享工作簿调用 `Save()`。严禁 `SaveAs`。
- 过程无论成功失败均通过 `Cleanup:` 恢复 `EnableEvents`/`ScreenUpdating` 等被改变的应用状态，并关闭本模块打开的工作簿；不挂单元格 `Change` 事件。

## 3. 涉及函数与模块

模块：`src/MdlExcel/`，依赖 `MdlInfra`、`MdlStorage`、`MdlSchema`；对上层不暴露 Excel 类型。`MdlSession` 分别调用 `MdlDiff` / `MdlMerge`，避免 Excel 适配层承担合并编排。心跳调度由本模块封装 `Application.OnTime`，回调目标为 `MdlSession.SessionHeartbeat`；会话层不直接访问 Excel 对象模型。

当前过程接口：

```text
Public Function OpenWorkbookHandle(ByVal path As String, ByVal role As WorkbookRole, ByRef handleId As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ListWorksheetNames(ByVal handleId As String, ByRef sheetNames As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ReadSheetSchema(ByVal handleId As String, ByVal sheetName As String, ByRef schema As SheetSchema, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ReadDeclaredSnapshot(ByVal handleId As String, ByRef targets As Variant, ByRef snapshot As SnapshotData, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ValidateEditableContents(ByVal handleId As String, ByRef targets As Variant, ByRef invalidCells As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Sub CopySnapshotSchema(ByRef snapshot As SnapshotData, ByRef expected As SessionSchema)
Public Function ValidateStructure(ByVal handleId As String, ByRef expected As SessionSchema, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function WriteDeclaredColumns(ByVal handleId As String, ByRef targets As Variant, ByRef mergedValues As Variant, ByRef mergedFonts As Variant, ByRef mergedFills As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function SaveWorkbook(ByVal handleId As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ScheduleSessionHeartbeat(ByVal intervalSeconds As Long, ByRef errorInfo As AppErrorInfo) As Boolean
Public Sub CancelSessionHeartbeat()
Public Sub CloseWorkbookHandle(ByVal handleId As String, ByVal saveChanges As Boolean)
Public Function RegisterCurrentTestWorkbook(ByRef handleId As String, ByRef errorInfo As AppErrorInfo) As Boolean
```

快照边界、内容守卫和句柄验证使用的关键私有过程及句柄类内部接口：

```text
Private Function ResolveSnapshotInputs(ByVal handleId As String, ByRef targets As Variant, ByRef workbook As Workbook, ByRef worksheet As Worksheet, ByRef schema As SheetSchema, ByRef declaredIndexes As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ValidateSheetColumns(ByRef worksheet As Worksheet, ByRef schema As SheetSchema, ByRef declaredIndexes As Variant, ByVal sheetName As String, ByRef invalidCells As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function FindLastDataRow(ByRef worksheet As Worksheet, ByRef anchorIndexes As Variant, ByRef declaredIndexes As Variant) As Long
Private Function ReadTextFontData(ByVal cell As Range, ByVal cellValue As Variant, ByRef data As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function CaptureFontSignature(ByVal fontObject As Object, ByRef signature As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ReadFillData(ByVal cell As Range, ByRef data As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function WriteTextFontData(ByVal cell As Range, ByVal cellValue As Variant, ByVal data As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ApplyFontSignature(ByVal fontObject As Object, ByVal signature As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function WriteFillData(ByVal cell As Range, ByVal data As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function IndexArraysEqual(ByRef leftValues As Variant, ByRef rightValues As Variant) As Boolean
Private Function AnchorMatricesEqual(ByRef leftValues As Variant, ByRef rightValues As Variant) As Boolean
Friend Sub CWorkbookHandle.Initialize(ByRef workbook As Workbook, ByVal owned As Boolean, ByVal role As WorkbookRole)
Friend Property Get CWorkbookHandle.WorkbookObject() As Workbook
Friend Property Get CWorkbookHandle.IsOwned() As Boolean
Friend Property Get CWorkbookHandle.Role() As WorkbookRole
Friend Sub CWorkbookHandle.Detach()
```

所有过程内部访问 Excel 对象；公共参数和返回值为字符串、标量、记录或数组。`MdlExcel` 通过私有 handle 类持有 workbook 引用；对共享路径要求 UNC，对本地副本要求位于 `%LOCALAPPDATA%`。字体 run 与填充均由确定形状的纯 Variant 记录表达，不把 Excel Font/Interior/Characters 对象泄漏到上层。写回按声明列处理并读回校验，应用状态恢复集中在 `Cleanup:`。

## 4. 功能范围

做：用宿主实例按会话选定路径打开/关闭、读取 schema/值/字符字体/填充快照、结构守卫、声明列值与格式写回、写后校验、`Save()`、区分可判定的共享路径/只读打开错误、心跳调度封装、错误清理。

不做：Excel 外部自动化实例、`SaveAs`、提交前备份、`UsedRange`/`Find` 边界搜索、结构编辑、非声明列写回、单元格事件监控、差异预览。

## 5. 测试方案

- Excel 集成测试放 `test/MdlExcel/` 并注明 Windows + 目标 Excel 版本前提；在未保存的测试工作簿上构造合成 sheet，不使用生产业务数据或调用 `SaveAs`。
- 心跳调度测试验证非法间隔被拒绝，以及有效回调可注册并立即取消，不等待定时任务触发。
- 覆盖本地/共享角色句柄、表头解析输入输出、值/字体 runs/填充快照往返、仅声明列写回、非声明列/锚点保持不变、锚点值或顺序变化拒写、结构变化拒写、字体颜色/粗体的局部字符段读写、填充读写、写入后格式逐项校验、保存失败及清理恢复。共享打开失败需覆盖不存在路径、不可达共享、访问拒绝、Workbook.ReadOnly 和损坏/不支持格式；验证错误类别、错误上下文及失败时不写入。只读状态不得被测试或提示认定为确定的锁冲突。
- 行边界探测覆盖中间存在整行空白后仍有数据、末行锚点为空但声明列有值、样式延伸到工作表末行但无值、锚点列末行不同及最大行边界；任何声明列中存在无法由锚点定位的非空数据时必须拒绝快照，不得静默漏行。
- 受控探测脚本 `tools/Probe-ExcelBoundary.ps1` 的 worker 硬超时为 60 秒，结束后最多等待 20 秒让隔离 Excel 退出，未退出则终止并报失败。它在 Excel 16.0 build 19127 的未保存内存工作簿上通过：A/C/E 被识别为锚点，F 的 product 列被排除；第 2、4、10 行有数据且中间 6 行全空时末行仍为 10；仅将格式延伸至第 1,048,576 行不扩大末行；第 30 行只有声明列值时检测为无锚点数据。该结果验证候选 API 行为，不代表真实业务模板实测。
- `test/MdlExcel/` 复用物理行/列矩阵验证 A/C/E 锚点、A–E 以外的 product 排除、内部空行、公式/非字符串拒绝、孤立 AMG 末行失败、锚点漂移拒写以及只写声明列。
- `tools/Probe-ExcelBoundary.ps1` 在 Excel 16.0 build 19127 的隔离未保存样例中复核 `End(xlUp)` 候选末行：锚点/声明列末行均为 10，row 2–10 中 6 行空白保留，工作表最大格式行不扩大末行；声明值位于 row 30 且没有锚点时被拒绝。报告位于 `docs/probes/anchor-boundary-2026-11-19.md`。
- 对非连续 AMG 列逐列写入，验证不会遗漏列；确认没有 `SaveAs`，并验证失败路径不留下打开句柄或关闭用户原有工作簿。
- 格式集成测试验证同一字符串内多个字体 run、字符位置和 run 边界、空值/零长度字符串的单元格级字体、插入/删除后的快照变化、直接填充属性往返、主题/自动颜色的稳定性，以及未声明列和非目标字符的格式不变。格式 API 性能需用代表性规模记录；不得对非声明列扫描字体 runs。
- 当前合成 Excel 测试已验证 A/C/E 锚点选择、内部空行及行边界、单元格内多个字体 run、颜色/粗体局部字体写回、直接纯色填充写回、写后原属性保持及锚点漂移拒绝；主题/自动颜色、图案填充、空字符串格式、代表性模板性能与真实 UNC 保存仍待 Phase 9。
- Excel 测试使用可控超时；本计划编写阶段不启动 VBA 编译/测试进程。

## 6. 风险与未决问题

- **代表性模板尚未在仓库中识别。** 已对两个准备期归档工作簿做脱敏 OpenXML 摘要；一个没有 AMG，另一个是 schema probe，均不能代表真实模板。隔离 Excel 探测支持 `End(xlUp)` 候选规则，但真实数据分布和性能仍未验证；若始终无法取得代表性模板，Phase 9 必须标记此项阻塞，不得宣称生产就绪。
- **行末边界与孤立声明值的处理。** 当前实现按锚点列和声明列分别确定末行并取最大值，保留 row 2 到末行的物理行位置；全空内部行保留为空行，声明列有值但该行所有锚点均空时失败关闭。隔离探测已验证该规则的边界行为，仍需代表性模板和最大规模数据验证性能。
- **用户可能在本地副本误改 product 锚点。** 暂定提交校验发现锚点相对 Base 改变即拒绝提交并要求撤销，不会写回锚点列；需确认是否应明确以中文提示到具体位置。
- 空白单元格按 Phase 1 的 `Empty` 哨兵规则处理；公式和非字符串非空值在签出/提交守卫中拒绝。需用目标 Excel 验证批量读取后 `Empty` 与零长度字符串的区分行为。
- Excel `Characters`/Font 格式读取可能需要逐段访问，格式 run 数量较大时会增加提交耗时；优先按连续相同格式 run 压缩并限制在声明列，Phase 9 记录代表性模板耗时。若主题色/自动色或复杂字符格式不能无损往返，暂定拒绝该单元格提交并保留本地会话，需在目标 Excel 实测后确认错误提示。
- 命名单元格样式或工作簿主题定义不作为全局对象同步；计划只保存并写回声明单元格实际读到的 Font/Interior 属性值，不保留命名样式引用。若业务要求保留样式身份或传播主题定义，需回来确认后另行扩展。
- **保存期间共享文件的其它访问行为及 `Save()` 失败后的结果可见性，须由 Phase 9 验证。** 不做提交前备份；结果不确定时由 Phase 7 保留会话和本地数据，不可直接清理。
- **Excel 的只读打开不一定由其他用户锁定造成。** 当前实现不单独识别 sharing violation；提示使用“无法取得写入访问，可能由占用、只读属性或权限导致”。若目标 Excel/Samba 后续提供稳定、可直接读取的锁冲突信号，再评估是否增加分类，不解析模糊描述来猜测。
- `Application` 事件处理器若后续确需增加，必须遵守 `AGENTS.md` §4.5 的错误处理和 `mInCommit` 次序要求。
