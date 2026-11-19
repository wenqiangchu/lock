# Phase 8：界面

**依赖：Phase 7。**

## 1. 阶段目的

提供普通用户可直接使用的 Ribbon 与窗体，使用户能选择 sheet/AMG 列、提交或放弃，并在发生单元格冲突时逐格查看 Base、Local、Remote 后选择保留 Local 或 Remote。

## 2. 实现方式

- Ribbon XML 定义三个按钮：“开始编辑”“提交”“放弃”。`MdlUI` 的 Ribbon 回调读取 `MdlSession.GetSessionView` 计算按钮状态，业务操作只调用 `MdlSession` 公共接口。
- `ShowTargetSelector` 使用 Office 自带 Microsoft Forms 运行时创建模态多选窗体，展示工作表、AMG 列和锁状态；确认后将目标表转换成 `BeginCheckout` 所需的三列矩阵。锁状态可能在确认前变化，签出仍按已确认策略整组获取；任一列竞争失败则整组取消。
- 冲突窗体一次显示一个单元格的定位及 Base/Local/Remote 三个只读、可滚动值框。用户选择 Local 或 Remote 后转到下一项；所有冲突均有选择后才启用“完成裁决”。完成后生成 `sheet/row/column/choice` 矩阵交给 Session。用户点取消或关闭窗体均调用 `CancelConflictReview`，不触发写回。
- `ResumeAfterConflictReview` 在写回前重新读取 Remote；若裁决期间冲突集合变化，UI 对新冲突集合重新显示并要求再次确认。直到冲突集合一致且所有选择完整前不调用写入。
- UI 不访问 `Workbook`/`Worksheet`/`Range`，不做业务判定，不自行释放锁或清理会话。展示数据由 Session 以纯数据结构提供。
- 用户可见文本使用中文；错误提示说明失败对象、是否保留本地会话及可执行的下一步，不显示未经确认的“提交成功”。

## 3. 涉及函数与模块

模块：`src/MdlUI/MdlUI.bas`、`CTargetSelectorEvents.cls`、`CConflictReviewEvents.cls` 与 `ribbon.xml`。业务层只依赖 `MdlSession`；界面使用 Office 内置 Ribbon 与 Microsoft Forms 类型，不引入第三方运行时。

当前过程接口：

```text
Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
Public Sub RibbonGetEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
Public Sub RibbonStartEdit(ByVal control As IRibbonControl)
Public Sub RibbonSubmit(ByVal control As IRibbonControl)
Public Sub RibbonAbandon(ByVal control As IRibbonControl)
Public Function IsRibbonButtonEnabled(ByVal buttonId As String, ByRef view As SessionView) As Boolean
Public Function BuildConflictChoiceMatrix(ByRef conflicts As Variant, ByRef selections As Variant, ByRef choices As Variant) As Boolean
Public Function FormatConflictValue(ByVal value As Variant) As String
Public Function ExcelColumnLabel(ByVal columnIndex As Long) As String
Private Function ShowTargetSelector(ByRef targets As Variant, ByRef selectedTargets As Variant) As Boolean
Private Function ShowConflictReview(ByRef conflicts As Variant, ByRef choices As Variant) As Boolean
Private Function AddValueBox(ByRef form As Object, ByVal name As String, ByVal labelText As String, ByVal left As Double, ByVal top As Double) As Object
Private Function AddButton(ByRef form As Object, ByVal name As String, ByVal caption As String, ByVal left As Double, ByVal top As Double, ByVal width As Double) As Object
Private Function AddLabel(ByRef form As Object, ByVal name As String, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal height As Double, ByVal caption As String) As Object
Public Sub CTargetSelectorEvents.Bind(ByVal formObject As Object, ByVal listBox As Object, ByVal acceptButton As Object, ByVal cancelButton As Object)
Public Property Get CTargetSelectorEvents.Accepted() As Boolean
Private Sub CTargetSelectorEvents.mAcceptButton_Click()
Private Sub CTargetSelectorEvents.mCancelButton_Click()
Public Sub CConflictReviewEvents.Bind(ByVal formObject As Object, ByRef conflicts As Variant, ByVal locationLabel As Object, ByVal countLabel As Object, ByVal baseValue As Object, ByVal localValue As Object, ByVal remoteValue As Object, ByVal previousButton As Object, ByVal nextButton As Object, ByVal localButton As Object, ByVal remoteButton As Object, ByVal finishButton As Object, ByVal cancelButton As Object)
Public Property Get CConflictReviewEvents.Confirmed() As Boolean
Public Property Get CConflictReviewEvents.ChoiceMatrix() As Variant
Private Sub CConflictReviewEvents.SelectChoice(ByVal choice As String)
Private Sub CConflictReviewEvents.RenderCurrent()
Private Function CConflictReviewEvents.AllChoicesMade() As Boolean
Private Function CConflictReviewEvents.ReadConflictBounds(ByRef conflicts As Variant, ByRef rowLow As Long, ByRef rowHigh As Long) As Boolean
Private Sub CConflictReviewEvents.mFinishButton_Click()
Private Sub CConflictReviewEvents.mCancelButton_Click()
Private Sub CConflictReviewEvents.mPreviousButton_Click()
Private Sub CConflictReviewEvents.mNextButton_Click()
Private Sub CConflictReviewEvents.mLocalButton_Click()
Private Sub CConflictReviewEvents.mRemoteButton_Click()
Private Sub CConflictReviewEvents.SelectChoice(ByVal choice As String)
Private Sub CConflictReviewEvents.RenderCurrent()
Private Function CConflictReviewEvents.SelectedCount() As Long
Private Function CConflictReviewEvents.ChoiceCaption(ByVal caption As String, ByVal selected As Boolean) As String
Private Function CConflictReviewEvents.AllChoicesMade() As Boolean
Private Function CConflictReviewEvents.ReadConflictBounds(ByRef conflicts As Variant, ByRef rowLow As Long, ByRef rowHigh As Long) As Boolean
```

窗体事件类持有 Office Forms 控件引用，不持有 Excel 业务对象。关闭窗体与“取消”使用相同的未提交裁决路径。

## 4. 功能范围

做：三个 Ribbon 按钮、sheet/AMG 多选器、进度/状态提示、逐格冲突裁决、放弃确认、会话状态对应的按钮启用/禁用。

不做：只读查看/谁在编辑面板、全量差异预览、自动选择冲突、跳过忙碌列的隐式行为、在线协作、超出 3 个按钮的 Ribbon 功能。

## 5. 测试方案

- `test/MdlUI/TestMdlUI.bas` 用内存数据覆盖 Ribbon 状态映射、冲突选择矩阵、未完成选择拒绝、空值/换行/长字符串显示格式及 A/E/AA 列地址转换，不读取 Excel 对象。
- Excel UI 集成测试验证开始编辑、提交、放弃按钮路径；冲突窗体每格只允许 Local/Remote，逐项选择后才完成，取消或关闭后共享文件未写且会话/列锁仍在。
- 当前已通过 Excel 16.0 build 19127 隔离项目编译和 MdlUI 内存测试；实际 Ribbon 装载、鼠标逐项选择、关闭窗体及共享文件未写入等 UI 集成路径仍需 Phase 9 验收。
- 覆盖空目标、重复目标、锁冲突提示、长值/空值显示、连续多个冲突、重复点击和窗体异常关闭。
- UI 测试使用专用工作簿并设置超时；计划编写阶段不启动编译进程。

## 6. 风险与未决问题

- **冲突很多时逐格窗体的导航和性能尚无数量上限。** 当前单格逐项呈现并提供上一项/下一项及已裁决计数，不增加全量差异预览；需在试用中确认是否易用。
- **Office Forms 运行时和 Ribbon 包装尚需实际 UI 实测。** `.xlam` 包结构已通过构建检查，但尚未在 Excel 用户界面实际装载并操作；若目标环境禁用 Forms 或 Ribbon XML 未正确装入，签出/裁决界面不可用，Phase 9 必须做加载验收。
- 冲突值按字符串原样显示，换行仅标准化显示；空值以“（空白）”标记。当前合同只允许字符串与 Empty，因此不提供日期、错误值或其他类型的编辑/裁决路径。
- 锁冲突时提示导致失败的 sheet/AMG 列；按已确认策略整组取消，不提供跳过占用列继续。
