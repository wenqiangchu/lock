# Phase 2：表结构解析

**依赖：Phase 1。**

## 1. 阶段目的

把工作表第 1 行表头解析为不依赖 Excel 的结构描述，识别可编辑 AMG 列、product 锚点列及列名指纹。完成后，签出和写回可以用一致的列定义校验目标，而不会由不同模块各自解释表头。

## 2. 实现方式

- `MdlExcel` 负责从工作簿读取表头并转成 `Variant` 数组；`MdlSchema` 只接收内存数组，不接触 `Workbook`、`Worksheet` 或 `Range`。
- 对表头做大小写无关的 AMG 前缀识别和规范化。规范化后重复的 AMG 列名立即报错，防止签出目标映射到错误列。
- 按已确认决策，product 列用于锚点且最多取 5 列；锚点唯一性属于业务前置条件，不在此模块检测。当前权威计划没有给出 product 列的具体表头识别规则，不把所有非 AMG 列自动视为 product；规则未明确时拒绝生成可签出的 schema。
- 返回表头顺序、列索引、列类别、选中 AMG 列索引、锚点列索引及稳定的列名指纹。指纹用于签出与提交时发现结构变化，不用于证明工作簿内容整体相同。
- 解析失败时拒绝该 sheet 的签出/提交，不生成部分可用的结构结果。

## 3. 涉及函数与模块

模块：`src/MdlSchema/`，只依赖 `MdlInfra`。

建议过程接口：

```text
Public Function ParseHeader(headerValues As Variant, ByRef schema As SheetSchema, ByRef err As AppErrorInfo) As Boolean
Public Function GetEditableColumns(schema As SheetSchema) As Variant
Public Function GetAnchorColumns(schema As SheetSchema) As Variant
Public Function BuildColumnFingerprint(schema As SheetSchema, columnIndexes As Variant) As String
Private Function NormalizeHeaderName(headerValue As Variant) As String
Private Function IsAmgHeader(normalizedName As String) As Boolean
```

`SheetSchema` 保存列序号、规范化表头、列类型、AMG 列集合、product/锚点列集合和结构指纹。其具体 VBA 类型布局在编码阶段按数组可测试性确定。

## 4. 功能范围

做：单行表头解析、按明确规则识别 AMG/product 列、重复 AMG 名报错、最多 5 个 product 锚点列的结构描述、列名指纹。

不做：打开工作簿、读取单元格对象、锚点值唯一性检测、按文件名区分 SDTM/ADaM/VLM/CT、自动修复或重命名表头、检查行列增删以外的格式差异。

## 5. 测试方案

- `test/MdlSchema/` 使用内存表头数组，不依赖 Excel。
- 覆盖只有 product 列、单个/多个 AMG 列、大小写混用、空表头、非文本值、规范化后 AMG 重名、无 product 锚点、超过 5 个锚点、单列与最大表头规模。
- 验证列索引保持原始顺序，指纹对相同结构稳定，表头顺序或名称变化会改变指纹。
- 解析错误必须可被调用方区分为输入错误，不得吞掉或返回不完整 schema。

## 6. 风险与未决问题

- **product 列的表头识别规则，以及多个 product 列如何组成锚点，当前权威计划未定义。** 选项包括已批准的表头前缀规则、明确配置的列名集合或模板映射。暂不照搬 `prepare/` 中的旧识别规则；在确定前，未知列不作为锚点且相关 sheet 不可签出。需回来确认分类规则与最多 5 列的组合方式。
- **指纹编码/哈希算法尚未指定。** 暂按有序、大小写规范化的表头序列构造稳定签名；不声称其具备密码学抗碰撞能力。需在 Phase 6 接入前确认是否需要哈希以及采用何种标准库外实现。
- 锚点唯一性按已确认决策不检测。若业务前置条件被破坏，行定位可能不可靠；遇到无法建立一致行映射时，调用方应失败关闭，不猜测写入位置。
