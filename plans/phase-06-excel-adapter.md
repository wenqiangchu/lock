# Phase 6：Excel 适配层

**依赖：Phase 2、Phase 3。** 与 Phase 5 可并行，双方通过 `MdlStorage` 接口契约衔接。

## 1. 阶段目的

建立唯一接触 Excel 对象模型的适配层。完成后，上层只使用会话句柄、schema 和内存数组即可打开/读取/写入/保存工作簿，不会持有或传递 `Workbook`、`Worksheet`、`Range` 对象。

## 2. 实现方式

- `MdlExcel` 使用宿主 Excel 实例，管理本地副本和共享主文件的打开句柄；句柄以内部 ID 暴露给上层，Excel 对象引用只保存在模块私有状态中。
- 共享主文件复制到 `%LOCALAPPDATA%` 工作区，由 `MdlStorage` 执行文件操作。Base 快照从本地副本在签出时读取；提交时重新读取共享文件形成 Remote。
- 读取表头后调用 `MdlSchema` 解析。用有界读取策略取得声明列、product 锚点值及目标行数组；按锚点关联 Base、Local、Remote 行，不以行号单独假定业务行相同。禁止 `Range.Find` 和 `UsedRange` 求边界。行边界方法需基于产品模板验证并记录。
- 提交前再次校验 sheet、表头/列名指纹、锚点列及其顺序/值映射、数组尺寸。若锚点被编辑、行增删/移动或结构不符，失败关闭，不进行写入。
- 写回只针对会话声明的 AMG 列，按列逐个写入，不以 `Union` 操作非连续区域；写入后校验目标数据，再对共享工作簿调用 `Save()`。严禁 `SaveAs`。
- 过程无论成功失败均通过 `Cleanup:` 恢复 `EnableEvents`/`ScreenUpdating` 等被改变的应用状态，并关闭本模块打开的工作簿；不挂单元格 `Change` 事件。

## 3. 涉及函数与模块

模块：`src/MdlExcel/`，依赖 `MdlInfra`、`MdlStorage`、`MdlSchema`；对上层不暴露 Excel 类型。`MdlSession` 分别调用 `MdlDiff` / `MdlMerge`，避免 Excel 适配层承担合并编排。

建议过程接口：

```text
Public Function OpenWorkbookHandle(path As String, role As WorkbookRole, ByRef handleId As String, ByRef err As AppErrorInfo) As Boolean
Public Function ReadSheetSchema(handleId As String, sheetName As String, ByRef schema As SheetSchema, ByRef err As AppErrorInfo) As Boolean
Public Function ReadDeclaredSnapshot(handleId As String, targets As Variant, ByRef snapshot As SnapshotData, ByRef err As AppErrorInfo) As Boolean
Public Function ValidateStructure(handleId As String, expected As SessionSchema, ByRef valid As Boolean, ByRef err As AppErrorInfo) As Boolean
Public Function WriteDeclaredColumns(handleId As String, targets As Variant, mergedValues As Variant, ByRef err As AppErrorInfo) As Boolean
Public Function SaveWorkbook(handleId As String, ByRef err As AppErrorInfo) As Boolean
Public Sub CloseWorkbookHandle(handleId As String, saveChanges As Boolean)
```

所有过程内部访问 Excel 对象；公共参数和返回值为字符串、标量、记录或数组。

## 4. 功能范围

做：用宿主实例打开/关闭、读取 schema/快照、结构守卫、声明列写回、写后校验、`Save()`、错误清理。

不做：Excel 外部自动化实例、`SaveAs`、提交前备份、`UsedRange`/`Find` 边界搜索、结构编辑、非声明列写回、单元格事件监控、差异预览。

## 5. 测试方案

- Excel 集成测试放 `test/MdlExcel/` 并注明 Windows + 目标 Excel 版本前提；使用专用可丢弃的合成工作簿，不使用生产业务数据。
- 覆盖本地/共享角色句柄、表头解析输入输出、快照往返、仅声明列写回、非声明列/锚点保持不变、锚点值或顺序变化拒写、结构变化拒写、写入后校验、保存失败、打开失败及清理恢复。
- 对非连续 AMG 列逐列写入，验证不会遗漏列；确认没有 `SaveAs`，并验证失败路径不留下打开句柄或关闭用户原有工作簿。
- Excel 测试使用可控超时；本计划编写阶段不启动 VBA 编译/测试进程。

## 6. 风险与未决问题

- **实际数据行边界规则尚未定稿。** 暂定仅使用锚点列与受限数组扫描/经验证的边界 API，发现空洞或尺寸歧义即拒绝操作；需用代表性模板确认是否存在尾部空锚点、内部空行及最大行数。
- **用户可能在本地副本误改 product 锚点。** 暂定提交校验发现锚点相对 Base 改变即拒绝提交并要求撤销，不会写回锚点列；需确认是否应明确以中文提示到具体位置。
- **值、公式、错误值和空单元格的读取/写回语义需与 Phase 3/4 共用。** 暂按值模式处理，不转换数据类型；需在集成测试确定公式单元格的准入规则。
- **保存期间共享文件的其它访问行为及 `Save()` 失败后的结果可见性，须由 Phase 9 验证。** 不做提交前备份；结果不确定时由 Phase 7 保留会话和本地数据，不可直接清理。
- `Application` 事件处理器若后续确需增加，必须遵守 `AGENTS.md` §4.5 的错误处理和 `mInCommit` 次序要求。
