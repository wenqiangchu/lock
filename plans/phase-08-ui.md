# Phase 8：界面

**依赖：Phase 7。**

## 1. 阶段目的

提供普通用户可直接使用的 Ribbon 与窗体，使用户能从配置的多份工作簿中选择一份，再选择其 sheet/AMG 列、提交或放弃，并在值/字体组件或填充组件发生冲突时逐项查看 Base、Local、Remote 后选择保留 Local 或 Remote。

## 2. 实现方式

- Ribbon XML 定义三个英文按钮：“Start Editing”“Submit”“Abandon”。`MdlUI` 的 Ribbon 回调读取 `MdlSession.GetSessionView` 计算按钮状态，业务操作只调用 `MdlSession` 公共接口。
- `ShowWorkbookSelector` 先展示本机配置中的 UNC/盘符绝对路径清单；用户选中一份后，`MdlSession` 只为该文件枚举工作表目标。`ShowTargetSelector` 再展示工作表、AMG 列和锁状态；确认后目标矩阵的路径必须全部等于所选路径。每次会话只操作一份工作簿。
- 配置文件由用户手动创建和维护；界面不提供编辑/保存配置的设置窗体。生产配置可填写 UNC 或映射到共享目录的绝对盘符路径；加载项不验证映射目标。配置缺失或格式错误时提示 `%LOCALAPPDATA%\MDDTLock\config.ini` 位置和修复指引。用户从清单选择的文件不存在、共享不可达或访问受限时，显示可判定的具体类别和所选路径；Excel 以只读方式打开时提示“无法取得写入访问”，说明可能由其他用户占用、只读属性或权限造成，不将其一概标成确定的锁冲突。
- 冲突窗体一次显示一个冲突组件的 sheet/row/column 和组件名（文字/字体或填充），以及 Base/Local/Remote 三方预览。文字/字体预览保留逐字符颜色、粗体等格式差异；填充预览展示各方填充状态。用户选择 Local 或 Remote 后转到下一项；所有组件冲突均有选择后才启用“完成裁决”。完成后生成 `sheet/row/column/component/choice` 矩阵交给 Session。用户点取消或关闭窗体均调用 `CancelConflictReview`，不触发写回。
- 所有模态窗体使用工程内的 `frmMddtDialog` UserForm 类承载运行时添加的 Forms 控件，通过 VBA `New` 实例化。不得通过 `CreateObject("Forms.UserForm.1")` 激活 UserForm；Office 16.0 实机探测返回错误 429，且会使选择器在开始前失败。构建和测试脚本先用 VBA 工程 API 创建空白 UserForm 组件，再导入模块；该容器不持有会话或 Excel 业务状态。
- `ResumeAfterConflictReview` 在写回前重新读取 Remote；若裁决期间冲突集合变化，UI 对新冲突集合重新显示并要求再次确认。直到冲突集合一致且所有选择完整前不调用写入。
- UI 不访问 `Workbook`/`Worksheet`/`Range`，不做业务判定，不自行释放锁或清理会话。展示数据由 Session 以纯数据结构提供。
- 用户可见文本使用英文；错误提示说明失败对象、已确认的原因类别、是否保留本地会话及可执行的下一步，不显示未经确认的提交成功信息。不能可靠区分的文件锁原因必须明确说原因未能确认。

## 3. 涉及函数与模块

模块：`src/MdlUI/MdlUI.bas`、`CTargetSelectorEvents.cls`、`CConflictReviewEvents.cls` 与 `ribbon.xml`。`tools/Build-Addin.ps1` 和 `tools/Run-VbaTests.ps1` 负责在隔离 VBA 工程中创建 `frmMddtDialog` UserForm 容器。业务层只依赖 `MdlSession`；界面使用 Office 内置 Ribbon 与 Microsoft Forms 类型，不引入第三方运行时。

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
Public Function FormatConflictComponent(ByVal componentName As String, ByVal value As Variant) As String
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

做：三个 Ribbon 按钮、配置工作簿单选器、所选文件的 sheet/AMG 多选器、进度/状态提示、按文字/字体与填充组件逐项冲突裁决、格式化预览、放弃确认、会话状态对应的按钮启用/禁用、明确的配置与文件打开失败提示。

不做：只读查看/谁在编辑面板、全量差异预览、自动选择冲突、跳过忙碌列的隐式行为、在线协作、超出 3 个按钮的 Ribbon 功能。

## 5. 测试方案

- Session 提供 N × 7 冲突矩阵 `[sheet, physicalRow, physicalColumn, component, base, local, remote]`；`BuildConflictChoiceMatrix` 只在所有选择有效时返回 N × 5 矩阵 `[sheet, physicalRow, physicalColumn, component, choice]`。`MdlSession` 将其按 sheet 分组为 Phase 4 的 N × 4 输入，不从预览文本重建数据。
- `FormatConflictComponent(componentName, value)` 对 `TextFont`/`Fill` 调用 `MdlInfra.FormatFormattingPreview`，返回规范化记录的可读摘要；TextFont 预览包括原字符串、字符 run 范围和对应字体属性，Fill 预览包括直接填充字段。它不修改记录，也不把预览摘要作为合并输入。
- `test/MdlUI/TestMdlUI.bas` 用内存数据覆盖 Ribbon 状态映射、含组件名的冲突选择矩阵、未完成选择拒绝、空值/换行/长字符串显示、字体 run 与填充预览格式化及 A/E/AA 列地址转换，不读取 Excel 对象。
- VBA 聚合测试在隔离 Excel 工程中实例化 `frmMddtDialog`，覆盖表单类创建且可由 VBA 实例化、可添加运行时控件，防止 `Forms.UserForm.1` 429 回归。
- Excel UI 集成测试验证从多个配置文件路径中选择单份工作簿、开始编辑、提交、放弃按钮路径；选中不存在/不可达/访问拒绝路径时显示对应错误且不取得锁；只读打开显示“无法取得写入访问”及可能原因。冲突窗体按文字/字体或填充组件提供 Local/Remote 选择，逐项选择后才完成，取消或关闭后共享文件未写且会话/列锁仍在。
- 实机冒烟先用本机临时配置显示工作簿选择窗体，再取消且不选择测试工作簿路径；分别验证配置缺失提示与选择器取消均不开始签出。临时配置和测试加载项在验收后移除。
- `.xlam` 包结构检查必须要求 `/customUI/customUI14.xml` 的 Content Type 为 `application/xml`；只检查 Ribbon 关系、部件存在和回调名称不足以证明 Excel 会呈现功能区。
- 2026-11-19 在 Office 2019 Version 2508 (Build 19127.20860) 对照验证：原包的 Ribbon 未出现；仅将临时副本的 Ribbon Content Type 改为 `application/xml` 后，Excel 显示 `MDDT Collaboration` 选项卡。随后通过 Excel 加载项 Browse 启用正式包 `MDDTLockIssue-ribbon-fix-20261119.xlam`，正式包也显示该选项卡；旧包和临时探测包未启用。
- 2026-11-20 回归：在本机交互 Excel 16.0 中加载修复包时未出现宏安全提示，Ribbon 标签及按钮可见；点击“Start Editing”后 `Select Workbook` 窗体正常显示测试配置中的 UNC 占位项，随后取消且未访问该路径。此前 `CreateObject("Forms.UserForm.1")` 返回错误 429；现由构建/测试脚本通过 VBA 工程 API 创建 `frmMddtDialog`，源码改用 `New frmMddtDialog`。定向 UI 测试和完整 VBA 套件通过，隔离测试 Excel 已退出。该验证只覆盖加载、选择窗体显示与取消，不覆盖目标选择、冲突裁决、共享签出/提交或真实模板。
- 覆盖空工作簿清单、单项/多项工作簿清单、取消工作簿选择、空目标、重复目标、锁冲突提示、文件不可用错误提示、长值/空值显示、连续多个冲突、重复点击和窗体异常关闭。
- UI 测试使用专用工作簿并设置超时；任何 Excel/VBA 自动化均须可控超时并确认隔离进程退出。

## 6. 风险与未决问题

- **冲突很多时逐格窗体的导航和性能尚无数量上限。** 当前单格逐项呈现并提供上一项/下一项及已裁决计数，不增加全量差异预览；需在试用中确认是否易用。
- **Office Forms 运行时和 Ribbon 操作仍需完整验收。** 已确认修复包可在本机 Excel 显示 Ribbon，选择工作簿窗体可显示并取消；旧 UserForm ProgID 方式会报错 429，现已由工程 API 创建的 `frmMddtDialog` 修复。目标选择窗体、完整冲突裁决窗体、远端写回与取消后的会话/锁状态仍需 Phase 9 实测。
- 冲突文本按原值显示，换行仅标准化；空值以“（空白）”标记。文字/字体冲突需要在预览中表达 run 边界和直接字体属性；若 Forms 控件不能保留局部格式，必须使用明确的文本/样式摘要预览，不得只显示丢失格式的纯文本并让用户误判。当前合同仍不提供日期、错误值或其它值类型的编辑/裁决路径。
- 锁冲突时提示导致失败的 sheet/AMG 列；按已确认策略整组取消，不提供跳过占用列继续。
- 配置维护固定为用户手工编辑本机 `%LOCALAPPDATA%\MDDTLock\config.ini`；条目可写 UNC 或映射到共享目录的绝对盘符路径，加载项不检查盘符映射关系。README 和配置缺失提示说明创建目录、创建/编辑文件、按 `WorkbookPath=` 增删条目以及版本 2 示例。不增加配置管理窗体。
