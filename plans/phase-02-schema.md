# Phase 2：表结构解析

**依赖：Phase 1。**

## 1. 阶段目的

把工作表第 1 行表头解析为不依赖 Excel 的结构描述，识别可编辑 AMG 列、product 锚点列及列名指纹。完成后，签出和写回可以用一致的列定义校验目标，而不会由不同模块各自解释表头。

## 2. 实现方式

- `MdlExcel` 负责从工作簿读取表头并转成 `Variant` 数组；`MdlSchema` 只接收内存数组，不接触 `Workbook`、`Worksheet` 或 `Range`。
- 对表头做大小写无关的 AMG 前缀识别和规范化。规范化后重复的 AMG 列名立即报错，防止签出目标映射到错误列。
- 非空且不以 `AMG` 开头的表头视为 product 列；空表头不分类为 product。锚点只从 A–E 中识别出的 product 列取得，按列序组成复合锚点；E 之后的 product 列仍为不可编辑列，但不参与锚点。A–E 至少要有一个可用 product 锚点，否则拒绝签出；非文本且非空的表头值若无法可靠分类，则拒绝建立 schema。锚点唯一性属于业务前置条件，不检测。
- 返回表头顺序、列索引、列类别、选中 AMG 列索引、锚点列索引及稳定的列名指纹。指纹用于签出与提交时发现结构变化，不用于证明工作簿内容整体相同。
- 解析失败时拒绝该 sheet 的签出/提交，不生成部分可用的结构结果。

## 3. 涉及函数与模块

模块：`src/MdlSchema/`，只依赖 `MdlInfra`。

当前过程接口：

```text
Public Function ParseHeader(ByRef headerValues As Variant, ByRef schema As SheetSchema, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function GetEditableColumns(ByRef schema As SheetSchema) As Variant
Public Function GetAnchorColumns(ByRef schema As SheetSchema) As Variant
Public Function BuildColumnFingerprint(ByRef schema As SheetSchema, ByVal columnIndexes As Variant) As String
Private Function IsAmgHeader(ByVal normalizedName As String) As Boolean
Private Function BuildFullFingerprint(ByRef schema As SheetSchema) As String
Private Sub ResetSchema(ByRef schema As SheetSchema)
```

`SheetSchema` 使用 `Public Type` 保存 `ColumnCount As Long`、`HeaderNames/NormalizedNames/ColumnKinds/AmgColumnIndexes/AnchorColumnIndexes As Variant` 和 `Fingerprint As String`。列清单均为按物理列序排列的一维 Variant 数组，`ColumnKinds` 与表头等长；空表头仍占据物理列位置，但不成为 product 或 AMG 列。`ParseHeader` 接受标量、1D 数组或单行/单列 2D Variant 数组并规范化为从物理列 1 开始的 schema。

## 4. 功能范围

做：单行表头解析、AMG/product 分类、从 A–E 选择 product 锚点、重复 AMG 名报错、最多 5 个锚点列的结构描述、列名指纹。

不做：打开工作簿、读取单元格对象、锚点值唯一性检测、按文件名区分 SDTM/ADaM/VLM/CT、自动修复或重命名表头、检查行列增删以外的格式差异。

## 5. 测试方案

- `test/MdlSchema/` 使用内存表头数组，不依赖 Excel。
- VBA 测试由 `tools/Run-VbaTests.ps1` 在受控的隐藏 Excel 进程中作为 VBA 执行宿主运行；`MdlSchema` 测试本身只传内存数组，不读取工作簿对象。runner 有硬超时，并在退出前确认新建 Excel 进程终止。
- 覆盖只有 product 列、单个/多个 AMG 列、大小写混用、非 AMG 列分类为 product、A–E 内外 product 列的锚点边界、空表头被忽略且不遮蔽其它有效锚点、非文本值、规范化后 AMG 重名、A–E 无 product 锚点、单列与最大表头规模。
- 验证列索引保持原始顺序，指纹对相同结构稳定，表头顺序或名称变化会改变指纹。
- 解析错误必须可被调用方区分为输入错误，不得吞掉或返回不完整 schema。
- 另用受控 Excel 样例探测第 1 行表头从 A 到 E 的列位置映射；探测脚本为 `tools/Probe-ExcelBoundary.ps1`，脱敏结构摘要脚本为 `tools/probe_xlsx_structure.py`，执行步骤与结果记录在 `docs/probes/anchor-boundary-2026-11-19.md`。报告只记录列号和分类，不输出实际表头文本。
- 归档样本结果：`prepare/测试结果.xlsx` 有 1 张表且无 AMG 表头；`prepare/_smoke/schema_probe.xlsx` 有 7 张表，其中 6 张含 AMG、1 张含 AMG 的样表没有 A–E product 锚点，另有样表在 E 后仍有 product 列。两者都未确认是代表性业务模板；这些结果只用于暴露待验证边界，不能据此改变 A–E 合同或声称真实模板符合前置条件。

## 6. 风险与未决问题

- 指纹当前采用带物理列号、列类别、规范化表头长度和值的有序长度前缀字符串，不使用哈希，也不声称具有密码学抗碰撞能力。超长表头可能导致指纹较长；当前工作簿规模下先接受此简单表示，若实际性能不满足再另行决策。
- 若 A–E 中没有非空且非 AMG 的 product 列，schema 校验失败并拒绝签出。隔离探针已验证列位规则，仍需 Phase 9 用代表性模板确认该前置条件覆盖全部目标 sheet；当前仓库没有可确认代表性的模板。
- 锚点唯一性按已确认决策不检测。若业务前置条件被破坏，行定位可能不可靠；遇到无法建立一致行映射时，调用方应失败关闭，不猜测写入位置。
