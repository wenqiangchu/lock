# Phase 8：界面

**依赖：Phase 7。**

## 1. 阶段目的

提供普通用户可直接使用的 Ribbon 与窗体，使用户能从配置的多份工作簿中选择一份，再选择其 sheet/AMG 列、提交或放弃，并在发生单元格冲突时逐格查看 Base、Local、Remote 后选择保留 Local 或 Remote。

## 2. 实现方式

- Ribbon XML 定义三个英文按钮：“Start Editing”“Submit”“Abandon”。`MdlUI` 的 Ribbon 回调读取 `MdlSession.GetSessionView` 计算按钮状态，业务操作只调用 `MdlSession` 公共接口。
- `ShowWorkbookSelector` 先展示本机配置中的 UNC 工作簿清单；用户选中一份后，`MdlSession` 只为该文件枚举工作表目标。`ShowTargetSelector` 再展示工作表、AMG 列和锁状态；确认后目标矩阵的路径必须全部等于所选路径。每次会话只操作一份工作簿。
- 配置文件由用户手动创建和维护；界面不提供编辑/保存配置的设置窗体。配置缺失或格式错误时提示 `%LOCALAPPDATA%\MDDTLock\config.ini` 位置和修复指引。用户从清单选择的文件不存在、共享不可达或访问受限时，显示可判定的具体类别和所选 UNC 路径；Excel 以只读方式打开时提示“无法取得写入访问”，说明可能由其他用户占用、只读属性或权限造成，不将其一概标成确定的锁冲突。
- 冲突窗体一次显示一个单元格的定位及 Base/Local/Remote 三个只读、可滚动值框。用户选择 Local 或 Remote 后转到下一项；所有冲突均有选择后才启用“完成裁决”。完成后生成 `sheet/row/column/choice` 矩阵交给 Session。用户点取消或关闭窗体均调用 `CancelConflictReview`，不触发写回。
- `ResumeAfterConflictReview` 在写回前重新读取 Remote；若裁决期间冲突集合变化，UI 对新冲突集合重新显示并要求再次确认。直到冲突集合一致且所有选择完整前不调用写入。
- UI 不访问 `Workbook`/`Worksheet`/`Range`，不做业务判定，不自行释放锁或清理会话。展示数据由 Session 以纯数据结构提供。
- 用户可见文本使用英文；错误提示说明失败对象、已确认的原因类别、是否保留本地会话及可执行的下一步，不显示未经确认的提交成功信息。不能可靠区分的文件锁原因必须明确说原因未能确认。

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
Private Function ShowWorkbookSelector(ByRef paths As Variant, ByRef selectedPath As String) As Boolean
Private Function ShowTargetSelector(ByRef targets As Variant, ByRef selectedTargets As Variant) As Boolean
Private Function ShowConflictReview(ByRef conflicts As Variant, ByRef choices As Variant) As Boolean
Private Function AddValueBox(ByRef form As Object, ByVal name As String, ByVal labelText As String, ByVal left As Double, ByVal top As Double) As Object
Private Function AddButton(ByRef form As Object, ByVal name As String, ByVal caption As String, ByVal left As Double, ByVal top As Double, ByVal width As Double) As Object
Private Function AddLabel(ByRef form As Object, ByVal name As String, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal height As Double, ByVal caption As String) As Object
Public Sub CTargetSelectorEvents.Bind(ByVal formObject As Object, ByVal listBox As Object, ByVal acceptButton As Object, ByVal cancelButton As Object, ByVal emptySelectionMessage As String)
Public Property Get CTargetSelectorEvents.Accepted() As Boolean
Public Property Get CTargetSelectorEvents.SelectedIndex() As Long
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

做：三个 Ribbon 按钮、配置工作簿单选器、所选文件的 sheet/AMG 多选器、进度/状态提示、逐格冲突裁决、放弃确认、会话状态对应的按钮启用/禁用、明确的配置与文件打开失败提示。

不做：只读查看/谁在编辑面板、全量差异预览、自动选择冲突、跳过忙碌列的隐式行为、在线协作、超出 3 个按钮的 Ribbon 功能。

## 5. 测试方案

- `test/MdlUI/TestMdlUI.bas` 用内存数据覆盖 Ribbon 状态映射、冲突选择矩阵、未完成选择拒绝、空值/换行/长字符串显示格式及 A/E/AA 列地址转换，不读取 Excel 对象。
- Excel UI 集成测试验证从多个配置文件路径中选择单份工作簿、开始编辑、提交、放弃按钮路径；选中不存在/不可达/访问拒绝路径时显示对应错误且不取得锁；只读打开显示“无法取得写入访问”及可能原因。冲突窗体每格只允许 Local/Remote，逐项选择后才完成，取消或关闭后共享文件未写且会话/列锁仍在。
- `.xlam` 包结构检查必须要求 `/customUI/customUI14.xml` 的 Content Type 为 `application/xml`；只检查 Ribbon 关系、部件存在和回调名称不足以证明 Excel 会呈现功能区。
- 2026-11-19 在 Office 2019 Version 2508 (Build 19127.20860) 对照验证：原包的 Ribbon 未出现；仅将临时副本的 Ribbon Content Type 改为 `application/xml` 后，Excel 显示 `MDDT Collaboration` 选项卡。随后通过 Excel 加载项 Browse 启用正式包 `MDDTLockIssue-ribbon-fix-20261119.xlam`，正式包也显示该选项卡；旧包和临时探测包未启用。
- 当前已通过 Excel 16.0 build 19127 隔离项目编译和 MdlUI 内存测试，正式包结构测试及标签加载通过。此冒烟测试仅确认加载与标签呈现；三个按钮的点击、窗体交互、共享文件未写入等 UI 集成路径仍需 Phase 9 验收。
- 覆盖空工作簿清单、单项/多项工作簿清单、取消工作簿选择、空目标、重复目标、锁冲突提示、文件不可用错误提示、长值/空值显示、连续多个冲突、重复点击和窗体异常关闭。
- UI 测试使用专用工作簿并设置超时；计划编写阶段不启动编译进程。

## 6. 风险与未决问题

- **冲突很多时逐格窗体的导航和性能尚无数量上限。** 当前单格逐项呈现并提供上一项/下一项及已裁决计数，不增加全量差异预览；需在试用中确认是否易用。
- **Office Forms 运行时和 Ribbon 操作尚未完整验收。** 正式包的 Ribbon 标签已在本机 Office 2019 中确认可加载，且 Content Type 必须为 `application/xml`；按钮点击、Forms 窗体以及目标环境策略仍需 Phase 9 实测。
- 冲突值按字符串原样显示，换行仅标准化显示；空值以“（空白）”标记。当前合同只允许字符串与 Empty，因此不提供日期、错误值或其他类型的编辑/裁决路径。
- 锁冲突时提示导致失败的 sheet/AMG 列；按已确认策略整组取消，不提供跳过占用列继续。
- 配置维护固定为用户手工编辑本机 `%LOCALAPPDATA%\MDDTLock\config.ini`，内容仅接受 UNC 路径；README 和配置缺失提示说明创建目录、创建/编辑文件、按 `WorkbookPath=` 增删条目以及版本 2 示例。不增加配置管理窗体。
