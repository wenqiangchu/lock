# Phase 4：三方合并

**依赖：Phase 3。**

## 1. 阶段目的

实现脱离 Excel 的 Base / Local / Remote 三方合并。完成后，单侧修改可自动合并，双方对同一单元格作出不同修改时会生成逐格冲突，并能应用用户选择的 Local 或 Remote 值。

## 2. 实现方式

- 输入为按 product 锚点对齐且结构一致的内存数组：签出快照 Base、本地当前值 Local、提交时读取的最新共享值 Remote，以及声明列集合。锚点列值用于确定同一业务行，不参与写回；锚点唯一性依赖业务前置条件，不重复检测。
- 对每个声明单元格按规则处理：Local=Base 时取 Remote；Remote=Base 时取 Local；Local=Remote 时取共同值；三者均不同则记录冲突并等待用户裁决。非声明列不参与结果写入集合。
- 冲突记录包含 sheet/行/列定位信息以及 Base、Local、Remote 三值；模块不生成 UI 字符串，也不持有 Excel 对象。
- 冲突裁决输入是每个冲突单元格的 Local/Remote 选择。任何冲突未选择或选择非法时，返回“未完成”，结果不可提交；全部选择完成后生成最终数组。
- 合并与裁决全程只在内存中进行。只有 `MdlSession` 确认全部裁决完成后，才可推进到写回阶段。

## 3. 涉及函数与模块

模块：`src/MdlMerge/`，只依赖 `MdlInfra`。

建议过程接口：

```text
Public Function MergeThreeWay(baseValues As Variant, localValues As Variant, remoteValues As Variant, declaredColumns As Variant, ByRef result As MergeResult, ByRef err As AppErrorInfo) As Boolean
Public Function ApplyConflictChoices(ByRef result As MergeResult, choices As Variant, ByRef err As AppErrorInfo) As Boolean
Public Function IsMergeResolved(result As MergeResult) As Boolean
Private Function ClassifyCell(baseValue As Variant, localValue As Variant, remoteValue As Variant) As MergeCellOutcome
```

`MergeResult` 包含合并值、冲突清单、待裁决数量和完成标志。公开接口只收发纯数据数组/记录。

## 4. 功能范围

做：三方逐格判定、冲突列表生成、逐格 Local/Remote 选择、未完成裁决检测。

不做：Excel 读写、窗体交互、全量差异预览、自动偏向任一方、自动重试或跨会话保存裁决、不声明列的写入。

## 5. 测试方案

- `test/MdlMerge/` 使用内存数组，验证无修改、仅 Local 修改、仅 Remote 修改、双方改成相同值、单格/多格冲突、不同列并发修改不冲突、空数组、单元素、最大数组和类型混合。
- 对冲突逐一选择 Local/Remote，检查最终值及未选择、重复选择、非法选项、错误索引均不能标记为完成。
- 验证输入维度不一致、锚点映射不一致或声明列越界时整体失败，不返回可写回的部分结果。
- 本阶段不启动 Excel；冲突窗体取消后共享文件未写入的集成验证放 Phase 8/9。

## 6. 风险与未决问题

- 值比较规则沿用 Phase 3 的待确认约定；两阶段必须共享同一 `ValuesEqual` 语义，否则会出现 Diff 与 Merge 判定不一致。暂定类型敏感、无隐式转换。
- **冲突 UI 的取消语义已确定：不写共享文件，保留本地会话和列锁。** 本模块只返回未完成裁决状态，不负责释放锁；需由 Phase 7 保证这个状态转移。
- 格式属性不属于当前范围；未来若提出格式同步，须先修订总计划并明确格式级冲突和写回规则，再调整本阶段数据结构。
