# Phase 8：界面

**依赖：Phase 7。**

## 1. 阶段目的

提供普通用户可直接使用的 Ribbon 与窗体，使用户能选择 sheet/AMG 列、提交或放弃，并在发生单元格冲突时逐格查看 Base、Local、Remote 后选择保留 Local 或 Remote。

## 2. 实现方式

- Ribbon 提供总计划指定的 3 个按钮：“开始编辑”“提交”“放弃”。按钮启用状态由 `MdlSession.GetSessionView` 决定，回调只调用 `MdlSession` 公共接口。
- 选择窗体展示可编辑 sheet、AMG 列和锁占用状态；用户确认后提交目标集合。发生锁冲突时按已确认策略整组取消并给出明确原因，不让用户误以为部分目标已签出。
- 冲突窗体一次呈现一个冲突单元格的定位及 Base/Local/Remote 值，用户必须选择 Local 或 Remote；显示裁决进度。完成所有选择才提交选择集合；取消/关闭窗体调用取消接口，不触发写回。
- UI 不访问 `Workbook`/`Worksheet`/`Range`，不做业务判定，不自行释放锁或清理会话。展示数据由 Session 以纯数据结构提供。
- 用户可见文本使用中文；错误提示说明失败对象、是否保留本地会话及可执行的下一步，不显示未经确认的“提交成功”。

## 3. 涉及函数与模块

模块：`src/MdlUI/`、Ribbon XML/回调定义；只依赖 `MdlSession`。

建议过程接口：

```text
Public Sub RibbonStartEdit(control As IRibbonControl)
Public Sub RibbonSubmit(control As IRibbonControl)
Public Sub RibbonAbandon(control As IRibbonControl)
Public Sub UpdateRibbonState(control As IRibbonControl, ByRef enabled As Variant)
Private Function ShowTargetSelector(currentView As SessionView) As TargetSelectionResult
Private Function ShowConflictReview(conflicts As Variant) As ConflictChoiceResult
```

窗体负责收集和返回选择结果，不持有 Excel 业务对象。关闭按钮与“取消”使用相同的未提交裁决路径。

## 4. 功能范围

做：三个 Ribbon 按钮、sheet/AMG 多选器、进度/状态提示、逐格冲突裁决、放弃确认、会话状态对应的按钮启用/禁用。

不做：只读查看/谁在编辑面板、全量差异预览、自动选择冲突、跳过忙碌列的隐式行为、在线协作、超出 3 个按钮的 Ribbon 功能。

## 5. 测试方案

- `test/MdlUI/` 覆盖状态到按钮 enabled 状态映射、选择结果传递、冲突逐项选择和取消结果；可用表单级测试替身验证，不需 Excel。
- Excel UI 集成测试验证开始编辑、提交、放弃按钮路径；冲突窗体每格只允许 Local/Remote，逐项选择后才完成，取消或关闭后共享文件未写且会话/列锁仍在。
- 覆盖空目标、重复目标、锁冲突提示、长值/空值显示、连续多个冲突、重复点击和窗体异常关闭。
- UI 测试使用专用工作簿并设置超时；计划编写阶段不启动编译进程。

## 6. 风险与未决问题

- **冲突很多时逐格窗体的导航和性能尚无数量上限。** 暂定单格逐项呈现并提供上一项/下一项及已裁决计数，不增加全量差异预览；需在试用中确认是否易用。
- **单个 Variant 值（错误值、长文本、日期）在窗体中的格式化规则未定。** 暂按不改变原始值，仅用安全的显示转换；选择只提交 Local/Remote 标记，不从文本反向解析值。
- 锁冲突时提示导致失败的 sheet/AMG 列；按已确认策略整组取消，不提供跳过占用列继续。
