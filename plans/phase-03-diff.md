# Phase 3：差异检测

**依赖：Phase 1。** Phase 2 与本阶段可并行。

## 1. 阶段目的

提供纯内存的逐单元格比较能力，从签出时 Base 与当前 Local 快照中产生本次编辑的变更集合。完成后，提交流程可以准确限定变更位置和声明列，供三方合并使用。

## 2. 实现方式

- `MdlExcel` 读取 Base/Local 的数据并传入 `MdlDiff`；本模块只处理数组和声明列，不读取 Excel 对象。快照矩阵第一维下标保留工作表物理行号，第二维下标为物理列号；空数据快照用 `Empty` 表示。
- 每次调用只处理一个 sheet；输入是行 × 物理列的 2D Variant Base/Local 矩阵与一维物理列索引数组。行号保留工作表物理行序，列索引沿用原始列号；行锚点对齐由调用方在进入本模块前完成。本模块输出空 `Empty`（无差异）或 N × 4 Variant 矩阵，列依次为物理行索引、物理列索引、Base、Local。工作表标识由上层目标上下文携带。
- 按单元格比较；相等则不产生记录，不同则产生记录。未声明列不进入比较结果，更不能由结果下游写回。
- 声明 AMG 列只接受字符串；非空单元格不是字符串时整体返回输入错误。公式由 `MdlExcel` 检查并拒绝。空单元格以 `Empty` 哨兵表示，`Empty` 只与 `Empty` 相等；字符串通过 `MdlInfra.CompareCellValues` 精确比较，不做大小写转换、去空格或类型强制转换。
- 若输入形状、声明列与数据维度不匹配，返回错误并不生成可提交的部分差异。

## 3. 涉及函数与模块

模块：`src/MdlDiff/`，只依赖 `MdlInfra`。

当前过程接口：

```text
Public Function CompareSnapshots(ByRef baseValues As Variant, ByRef localValues As Variant, ByRef declaredColumns As Variant, ByRef changes As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
```

差异记录是纯数据，不包含 Range 引用；错误时输出 `Empty`，不得返回部分差异。

## 4. 功能范围

做：声明列内的逐格值差异、变更集合生成、输入维度校验、内存单元测试。

不做：访问 Excel、用户界面、共享文件读取、三方冲突判定、写回、增删行列或排序检测/修复。不得扩大到未声明列。

## 5. 测试方案

- `test/MdlDiff/`：相同快照、单格变化、多格变化、空集合、单列/单行、最大数组、字符串大小写/空格差异、`Empty` 与 `""`、非字符串值拒绝、维度不一致、声明列缺失或越界。
- VBA 测试由 `tools/Run-VbaTests.ps1` 在受控的隐藏 Excel 进程中作为 VBA 执行宿主运行；`MdlDiff` 过程只接触内存数组，不访问 Excel 对象模型。
- 验证产物只包含声明列，变更记录中的 Base/Local 与输入保持一致；相等值不产生伪差异。
- 测试纯内存运行，不依赖 Excel。若需要验证从工作簿读取的数组形状，在 Phase 6 的 Excel 集成测试另测。

## 6. 风险与未决问题

- 字符串比较由 `MdlInfra.CompareCellValues` 统一执行；当前精确区分大小写、空格和 `Empty`/`""`。若业务认为这些文本差异不重要，需先更改统一规则再实现。
- 行对齐依赖稳定 product 锚点和不改结构/不排序的合同；锚点唯一性不检测。Local 锚点与 Base 不匹配时由调用方失败关闭，不能按行号猜测对应关系。
