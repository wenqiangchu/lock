# Phase 7：会话状态机

**依赖：Phase 4、Phase 5、Phase 6。**

## 1. 阶段目的

把 schema、差异、合并、锁、存储与 Excel 适配层编排成完整签出、编辑、提交、冲突裁决和放弃流程。完成后，即使 UI 仅由临时按钮触发，也能安全跑通一次会话。

## 2. 实现方式

- `MdlSession` 维护 `Idle → TargetSelected → CheckingOut → CheckedOut` 与 `CheckedOut → Committing ↔ ConflictReview → Committing → Idle` 状态；取消裁决或提交失败回到 `CheckedOut`，放弃经过 `Aborting` 回到 `Idle`。
- 签出时校验目标 schema，按整组取得列锁，创建本地工作区/副本，打开本地副本并读取 Base。任一步骤失败时按相反顺序回滚本次资源，不形成半个活动会话。
- 配置允许登记多份 UNC 工作簿，但每次会话必须先选定且只绑定一份路径。`BeginCheckout` 校验所有目标都属于该路径，并将其固定在会话状态中；配置文件之后被修改也不改变活动会话的提交目标。
- 编辑时不访问共享文件、不挂单元格 `Change` 事件；周期性续约锁。续约失败或 fencing 校验失败时标记会话不可提交，但保留本地文件供用户恢复。
- `SubmitChanges` 校验活动会话与全部列锁后取得共享提交门闩，打开 Remote 并逐 sheet 读取 Local/Remote 的值、逐字符字体和填充快照；对 Base/Local 调用 `MdlDiff`，再由 `MdlMerge` 分别合并 `TextFont` 与 `Fill` 组件。冲突通过纯 Variant 矩阵返回，列为 `[sheet, physicalRow, physicalColumn, component, base, local, remote]`，不暴露 Excel 对象。
- 初次发现冲突时关闭 Remote 且释放提交门闩，保留本地副本、会话和列锁并进入 `ConflictReview`。UI 的选择矩阵为 `[sheet, physicalRow, physicalColumn, component, "LOCAL"|"REMOTE"]`。恢复裁决时重新读取 Remote 并重新合并；只有冲突坐标、组件及 Base/Local/Remote 数据均与上次展示完全相同，才接受选择。冲突集合或数据已变化时返回新列表并继续等待裁决，不写共享文件。
- 用户取消裁决时不得写共享工作簿，回到 `CheckedOut`，保留本地会话和列锁。全部冲突解决后才允许声明列的值、字体和填充写回并调用 `Save()`；成功确认后关闭句柄、释放列锁并清理本地工作区。
- 写入前和 `Save()` 前再次校验全部列租约及提交门闩。值和格式写入均限于 sheet/声明列；保存异常一律标记为结果不确定，保留本地工作区与锁且禁止盲目重试。提交成功但本地清理或释放锁失败时，向调用方返回“提交完成但清理有警告”，不得把清理错误伪装成提交失败或完全成功。
- 放弃前由 UI 确认丢弃；关闭本地副本不保存，释放锁并清理工作区。各清理失败要记录具体未完成项，不得报告为完全成功。

## 3. 涉及函数与模块

模块：`src/MdlSession/`，按 `AGENTS.md` 可协调各底层模块；不直接访问 Excel 对象。续约决策由会话状态机负责，定时回调由 `MdlExcel.ScheduleSessionHeartbeat` 注册、`MdlExcel.CancelSessionHeartbeat` 取消。

当前过程接口：

```text
Public Function GetConfiguredWorkbookPaths(ByRef paths As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function GetAvailableTargets(ByVal workbookPath As String, ByRef targets As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function BeginCheckout(ByVal workbookPath As String, ByRef targets As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function SubmitChanges(ByRef conflicts As Variant, ByRef errorInfo As AppErrorInfo) As SubmitOutcome
Public Function ResumeAfterConflictReview(ByRef choices As Variant, ByRef conflicts As Variant, ByRef errorInfo As AppErrorInfo) As SubmitOutcome
Public Function CancelConflictReview() As Boolean
Public Function AbandonSession(ByVal userConfirmed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function GetSessionView() As SessionView
Public Sub SessionHeartbeat()
Public Function IsSessionTransitionAllowed(ByVal currentState As SessionState, ByVal nextState As SessionState) As Boolean
Private Sub ScheduleHeartbeat()
Private Sub CancelHeartbeat()
Private Function TransitionTo(ByVal nextState As SessionState, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ExecuteCommit(ByRef choices As Variant, ByVal isConflictReview As Boolean, ByRef conflicts As Variant, ByRef errorInfo As AppErrorInfo) As SubmitOutcome
Private Function BuildCurrentMerges(ByVal remoteHandle As String, ByRef merges() As MergeResult, ByRef conflicts As Variant, ByRef hasLocalChanges As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ValidateSessionLeases(ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ValidateCommitAuthority(ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ConflictViewsEqual(ByRef leftConflicts As Variant, ByRef rightConflicts As Variant) As Boolean
Private Function ApplySessionConflictChoices(ByRef merges() As MergeResult, ByRef choices As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Private Sub RollbackCheckout(ByRef errorInfo As AppErrorInfo)
Private Function ReleaseSessionResources(ByRef errorInfo As AppErrorInfo) As Boolean
Private Sub ResetSession()
```

目标输入归一化与多 sheet 分组由以下私有过程完成：

```text
Private Function NormalizeTargets(ByRef targets As Variant, ByVal expectedFile As String, ByRef normalizedTargets As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function BuildTargetGroups(ByRef targets As Variant, ByRef sheetNames As Variant, ByRef groupTargets As Variant, ByRef groupCount As Long, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function FindSheetIndex(ByRef names() As String, ByVal groupCount As Long, ByVal sheetName As String) As Long
```

会话数据至少保存 session ID、状态、目标列清单、Base 值/字体/填充快照、结构指纹、锁 lease、工作区路径和本地句柄 ID；远端句柄仅在提交门闩持有期间短暂存在。冲突矩阵携带组件类型和原始纯 Variant 数据，UI 只提交冲突坐标、组件名和 Local/Remote 标签，不从显示文本反向解析数据。不得把 Excel 对象传给 `MdlUI`。

`GetConfiguredWorkbookPaths` 返回配置中的有序 UNC 路径列表。`GetAvailableTargets` 返回 N×4 纯 Variant 矩阵 `[UNC, sheet, AMG header, lock status]`，只枚举所选工作簿中可解析且含 AMG 列的 sheet。锁状态是只读提示，未知状态按不可确认展示；真正签出仍由整组锁获取决定。一个目标矩阵不得混入不同工作簿路径。

## 4. 功能范围

做：完整生命周期编排、从多路径配置选择单一会话工作簿、合法状态迁移、列锁续约/校验、签出回滚、值/字体/填充冲突取消保留、失败时保留本地数据、成功/放弃清理。

不做：多进程会话恢复 UI、实时协同、自动解决冲突、结构编辑、静默忽略未声明列改动、自动重试不确定结果的保存、全量差异预览。

## 5. 测试方案

- `test/MdlSession/` 对合法/非法状态迁移、Idle 状态拒绝提交/放弃和 `SessionView` 按状态计算做内存级测试；Excel 集成测试覆盖从至少两条配置路径中选择一份、会话路径固定、拒绝混合路径目标、签出→修改值/字体/填充→无冲突提交、无本地改动不写共享、值/字体/填充冲突进入裁决、取消裁决不写共享且保留本地和锁、Remote 在裁决期间变化后重显冲突、裁决完成后仅声明列值与格式写回、提交失败保留数据、放弃确认与取消。
- 覆盖重复点击、状态中提交/放弃、结构变化、锁过期、写入/保存/清理失败；Excel 进程和测试均必须受控超时。
- 当前 `test/MdlSession/TestMdlSession.bas` 仅覆盖状态迁移、Idle 操作拒绝和 `SessionView`；完整 UNC 签出/提交场景尚未实现为可运行测试，需依赖 Phase 9 的真实共享环境验收。
- 本地 Excel 集成 runner 不等价于 UNC/Samba 多客户端验收；无法连接受控共享时记录为 Phase 9 阻塞，不以模拟测试替代。

## 6. 风险与未决问题

- **冲突裁决期间 Remote 可能继续变化。** 暂定裁决结束后重新读取 Remote，并基于原 Base/最新 Local/新 Remote 再合并；若冲突集合变化则重新展示，不直接套用过期选择。需确认该重复裁决行为的用户体验。
- **保存报错不一定等价于共享文件未改变。** 暂按结果不确定处理：保留本地工作区与列锁、禁用再次提交，不自动重试；当前没有安全的自动恢复入口，用户需核对共享文件后明确放弃或由后续版本恢复。Phase 9 需在真实 SMB 上验证可观察结果和操作指引。
- **本地崩溃后的会话恢复尚未实现。** 当前状态只保存在 Excel 进程内存中；Excel/进程崩溃后不恢复会话或自动清理工作区，锁依赖租约过期。工作区保留策略与人工恢复入口仍待确认，Phase 9 需记录残留文件和锁过期行为。
- 多列锁冲突按已确认的整组取消执行。不同列同时提交通过 Phase 5 的短期提交门闩串行处理，避免共享工作簿同时保存；门闩只覆盖共享访问阶段，不阻止并行本地编辑。
- 多工作簿配置只影响可选目标；一个会话仍只绑定一份工作簿。需测试文件路径被错误混入目标矩阵时在取得任何锁之前整体拒绝。
