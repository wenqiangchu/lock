# Phase 4：三方合并

**依赖：Phase 3。**

## 1. 阶段目的

实现脱离 Excel 的 Base / Local / Remote 三方合并，覆盖声明单元格的文字/逐字符字体组件与直接填充组件。完成后，单侧组件修改可自动合并，双方对同一组件作出不同修改时会生成冲突，并能应用用户选择的 Local 或 Remote 组件。

## 2. 实现方式

- 输入为按 product 锚点对齐且结构一致的纯数据快照：签出快照 Base、本地当前 Local、提交时读取的最新共享 Remote，以及声明列集合。锚点列值用于确定同一业务行，不参与写回；锚点唯一性依赖业务前置条件，不重复检测。
- 每次调用只处理一个 sheet；值矩阵、逐字符字体映射矩阵与单元格填充矩阵均按物理行号/列号索引，空数据集用 `Empty` 表示。输出包含合并后的三类数据、N × 6 冲突矩阵（物理行、物理列、组件名、Base、Local、Remote）、冲突总数、待裁决数和完成标志。
- 每个声明单元格分别合并两个组件：`TextFont`（精确字符串及与字符位置绑定的直接字体属性映射）和 `Fill`（整格直接填充）。每个组件分别使用 Base/Local/Remote 三方规则：Local=Base 取 Remote；Remote=Base 取 Local；Local=Remote 取共同值；三者均不同则生成该组件冲突。不同组件可自动组合，例如 Local 改字体而 Remote 改填充时无冲突；双方改同一组件为不同结果时需裁决。非声明列不参与比较或写回。
- 字符串值与字符字体映射必须绑定为一个 `TextFont` 组件整体比较和裁决，不逐字符拼接双方结果。这样避免文字长度/位置改变后把一方格式套到另一方不同字符上；代价是双方改了不同字符的字体也会产生组件级冲突。填充独立合并，减少与字体变化无关的冲突。
- 合并模块冲突记录包含物理行/列、组件名及 Base/Local/Remote 纯数据；工作表名由 MdlSession 聚合时补入。模块不生成 UI 字符串，也不持有 Excel 对象。
- 冲突裁决输入为每项冲突的 Local/Remote 选择。任何组件冲突未选择或选择非法时，返回“未完成”，结果不可提交；全部选择完成后生成最终组件数据。
- 选择输入为 N × 4 Variant 矩阵，列依次为物理行、物理列、组件名和 Local/Remote。应用前先校验数量、坐标、重复项、组件名和选项；任何缺选或非法选择均不应用部分裁决。
- 合并与裁决全程只在内存中进行。只有 `MdlSession` 确认全部裁决完成后，才可推进到写回阶段。

## 3. 涉及函数与模块

模块：`src/MdlMerge/`，只依赖 `MdlInfra`。

当前过程接口：

```text
Public Function MergeThreeWay(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef remoteValues As Variant, ByRef baseFonts As Variant, ByRef localFonts As Variant, ByRef remoteFonts As Variant, ByRef baseFills As Variant, ByRef localFills As Variant, ByRef remoteFills As Variant, ByRef declaredColumns As Variant, ByRef result As MergeResult, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ApplyConflictChoices(ByRef result As MergeResult, ByRef choices As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function IsMergeResolved(ByRef result As MergeResult) As Boolean
Private Function ClassifyComponent(ByVal componentName As String, ByRef baseComponent As Variant, ByRef localComponent As Variant, ByRef remoteComponent As Variant, ByRef errorInfo As AppErrorInfo) As MergeCellOutcome
Private Sub ResetMergeResult(ByRef result As MergeResult)
```

`MergeResult` 包含 `MergedValues`、`MergedFonts`、`MergedFills`，组件级冲突清单、待裁决数量和完成标志。组件 `TextFont` 的冲突数据由字符串值与对应字体 run 组成；`Fill` 冲突数据是直接填充状态。公开接口只收发纯数据数组/记录，不传递 Excel 对象。

## 4. 功能范围

做：值/字体组件和填充组件的三方合并、组件级冲突列表、逐项 Local/Remote 选择、未完成裁决检测。

不做：Excel 读写、窗体交互、全量差异预览、自动偏向任一方、自动重试或跨会话保存裁决、不声明列的写入。

## 5. 测试方案

- `test/MdlMerge/` 使用内存数组，验证文字变化、字符颜色/粗体变化、填充变化、Local 改字体与 Remote 改填充自动组合、同组件相同修改、同组件不同修改冲突、文字长度变化时不跨快照拼接字符格式、不同列并发修改、空数组、单元素、最大数组和非法格式数据。
- 对每个文字/字体或填充冲突分别选择 Local/Remote，检查最终组件数据；未选择、重复选择、非法组件、非法选项或错误索引均不能标记为完成。
- 验证输入维度不一致、锚点映射不一致或声明列越界时整体失败，不返回可写回的部分结果。
- 本阶段不启动 Excel；冲突窗体取消后共享文件未写入的集成验证放 Phase 8/9。
- VBA 测试由 tools/Run-VbaTests.ps1 在受控隐藏 Excel 进程中作为执行宿主运行；合并实现只处理内存矩阵，不访问 Excel 对象模型。

## 6. 风险与未决问题

- 字符串与空白比较沿用 Phase 1 的 `MdlInfra.CompareCellValues`，确保 Diff/Merge 判定一致；非空非字符串和公式必须在 Excel 适配阶段拒绝。
- **冲突 UI 的取消语义已确定：不写共享文件，保留本地会话和列锁。** 本模块只返回未完成裁决状态，不负责释放锁；需由 Phase 7 保证这个状态转移。
- 逐字符格式映射的字符索引依赖字符串内容和长度；本阶段以 `TextFont` 整体为合并单元，不进行跨版本字符对齐或格式拼接。若未来需要细粒度字符级合并，须先另行设计并测试插入、删除和 Unicode 字符映射。
