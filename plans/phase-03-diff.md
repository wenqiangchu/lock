# Phase 3：差异检测

**依赖：Phase 1。** Phase 2 与本阶段可并行。

## 1. 阶段目的

提供纯内存比较能力，从签出时 Base 与当前 Local 快照中产生本次编辑的变更集合。比较对象包括声明 AMG 单元格的字符串值与逐字符直接字体格式，以及单元格直接填充；完成后，提交流程可以准确限定变更位置和声明列，供三方合并使用。

## 2. 实现方式

- `MdlExcel` 读取 Base/Local 的值、字体映射和填充并传入 `MdlDiff`；本模块只处理纯 Variant 数据和声明列，不读取 Excel 对象。值矩阵、字体映射矩阵与填充矩阵均以物理行号和物理列号索引；空数据快照用 `Empty` 表示。
- 每次调用只处理一个 sheet；行锚点对齐由调用方在进入本模块前完成。本模块输出空 `Empty`（无差异）或 N × 5 Variant 矩阵，列依次为物理行、物理列、组件名（`TextFont` 或 `Fill`）、Base 组件数据、Local 组件数据。同一单元格的文字/字体和填充可分别产生记录；工作表标识由上层目标上下文携带。
- `TextFont` 组件由精确字符串值和按字符位置记录的直接字体属性映射构成；文字发生插入/删除时，文本与字体映射作为整体比较。`Fill` 组件表示整格直接填充。任一组件相等则不产生该组件记录，不同则产生记录。
- 只有声明 AMG 列参与比较；非声明列、product 锚点及其它列不进入差异结果，也不能由下游写回。字体/填充属性仅在声明列中比较。
- 声明 AMG 列的非空值只接受字符串；公式由 `MdlExcel` 检查并拒绝。空单元格以 `Empty` 哨兵表示，`Empty` 只与 `Empty` 相等；字符串通过 `MdlInfra.CompareCellValues` 精确比较，不做大小写转换、去空格或类型强制转换。格式映射由 `MdlInfra.CompareFormattingData` 比较规范化纯数据。
- 若输入形状、声明列与数据维度不匹配，返回错误并不生成可提交的部分差异。

## 3. 涉及函数与模块

模块：`src/MdlDiff/`，只依赖 `MdlInfra`。

当前过程接口：

```text
Public Function CompareSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef baseFonts As Variant, ByRef localFonts As Variant, ByRef baseFills As Variant, ByRef localFills As Variant, ByRef declaredColumns As Variant, ByRef changes As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
```

差异记录是纯数据，不包含 Range 引用；文字/字体作为一个比较组件，填充为独立组件。错误时输出 `Empty`，不得返回部分差异。

## 4. 功能范围

做：声明列内逐格值、逐字符字体映射和直接填充差异，变更集合生成，输入维度校验及内存单元测试。

不做：访问 Excel、用户界面、共享文件读取、三方冲突判定、写回、增删行列或排序检测/修复。不得扩大到未声明列。

## 5. 测试方案

- `test/MdlDiff/`：相同快照、只改值、只改一个字符颜色/粗体、不同字符段格式变化、空白/零长度字符串的单元格级字体变化、文本插入/删除导致字体映射变化、只改填充、文字/字体与填充分离变化、多个组件同格变化、空集合、单列/单行、最大数组、字符串大小写/空格差异、`Empty` 与 `""`、非字符串值拒绝、矩阵维度不一致、声明列缺失或越界。
- VBA 测试由 `tools/Run-VbaTests.ps1` 在受控的隐藏 Excel 进程中作为 VBA 执行宿主运行；`MdlDiff` 过程只接触内存数组，不访问 Excel 对象模型。
- 验证产物只包含声明列，组件类型与 Base/Local 数据保持原样；未变化的组件不产生伪差异，值/字体及填充可分别报告。
- 测试纯内存运行，不依赖 Excel。若需要验证从工作簿读取的数组形状，在 Phase 6 的 Excel 集成测试另测。

## 6. 风险与未决问题

- 字符串比较由 `MdlInfra.CompareCellValues` 统一执行；当前精确区分大小写、空格和 `Empty`/`""`。若业务认为这些文本差异不重要，需先更改统一规则再实现。
- 逐字符字体数据必须与字符串长度及字符顺序一致；长度不匹配、无法读取或无法规范化的字体数据整体报错，不能把字体变化误当作无差异。
- 行对齐依赖稳定 product 锚点和不改结构/不排序的合同；锚点唯一性不检测。Local 锚点与 Base 不匹配时由调用方失败关闭，不能按行号猜测对应关系。
