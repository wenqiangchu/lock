# Phase 3：差异检测

**依赖：Phase 1。** Phase 2 与本阶段可并行。

## 1. 阶段目的

提供纯内存的逐单元格比较能力，从签出时 Base 与当前 Local 快照中产生本次编辑的变更集合。完成后，提交流程可以准确限定变更位置和声明列，供三方合并使用。

## 2. 实现方式

- `MdlExcel` 读取 Base/Local 的数据并传入 `MdlDiff`；本模块只处理数组和声明列，不读取 Excel 对象。
- 输入包含工作表/列标识、按 product 锚点关联的 Base 与 Local 行、目标列数组，以及本次声明列集合。输出为变更记录数组，每项含目标列、锚点对应的行索引、Base 值、Local 值和变化类别；锚点值只用于对齐，不写回。
- 按单元格比较；相等则不产生记录，不同则产生记录。未声明列不进入比较结果，更不能由结果下游写回。
- 声明 AMG 列只接受字符串；非空单元格不是字符串时整体返回输入错误。公式由 `MdlExcel` 检查并拒绝。空单元格以 `Empty` 哨兵表示，`Empty` 只与 `Empty` 相等；字符串通过 `MdlInfra.CompareCellValues` 精确比较，不做大小写转换、去空格或类型强制转换。
- 若输入形状、声明列与数据维度不匹配，返回错误并不生成可提交的部分差异。

## 3. 涉及函数与模块

模块：`src/MdlDiff/`，只依赖 `MdlInfra`。

建议过程接口：

```text
Public Function CompareSnapshots(baseValues As Variant, localValues As Variant, declaredColumns As Variant, ByRef changes As Variant, ByRef err As AppErrorInfo) As Boolean
Private Function MakeChangeRecord(sheetKey As String, rowIndex As Long, columnIndex As Long, baseValue As Variant, localValue As Variant) As ChangeRecord
```

`ChangeRecord` 是纯数据，不包含 Range 引用；行地址可由上层结合 sheet 与锚点信息展示。

## 4. 功能范围

做：声明列内的逐格值差异、变更集合生成、输入维度校验、内存单元测试。

不做：访问 Excel、用户界面、共享文件读取、三方冲突判定、写回、增删行列或排序检测/修复。不得扩大到未声明列。

## 5. 测试方案

- `test/MdlDiff/`：相同快照、单格变化、多格变化、空集合、单列/单行、最大数组、字符串大小写/空格差异、`Empty` 与 `""`、非字符串值拒绝、维度不一致、声明列缺失或越界。
- 验证产物只包含声明列，变更记录中的 Base/Local 与输入保持一致；相等值不产生伪差异。
- 测试纯内存运行，不依赖 Excel。若需要验证从工作簿读取的数组形状，在 Phase 6 的 Excel 集成测试另测。

## 6. 风险与未决问题

- 字符串比较由 `MdlInfra.CompareCellValues` 统一执行；当前精确区分大小写、空格和 `Empty`/`""`。若业务认为这些文本差异不重要，需先更改统一规则再实现。
- 行对齐依赖稳定 product 锚点和不改结构/不排序的合同；锚点唯一性不检测。Local 锚点与 Base 不匹配时由调用方失败关闭，不能按行号猜测对应关系。
