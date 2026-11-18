# Phase 7：会话状态机

**依赖：Phase 4、Phase 5、Phase 6。**

## 1. 阶段目的

把 schema、差异、合并、锁、存储与 Excel 适配层编排成完整签出、编辑、提交、冲突裁决和放弃流程。完成后，即使 UI 仅由临时按钮触发，也能安全跑通一次会话。

## 2. 实现方式

- `MdlSession` 维护 `Idle → TargetSelected → CheckingOut → CheckedOut` 与 `CheckedOut → Committing ↔ ConflictReview → Committing → Idle` 状态；取消裁决或提交失败回到 `CheckedOut`，放弃经过 `Aborting` 回到 `Idle`。
- 签出时校验目标 schema，按整组取得列锁，创建本地工作区/副本，打开本地副本并读取 Base。任一步骤失败时按相反顺序回滚本次资源，不形成半个活动会话。
- 编辑时不访问共享文件、不挂单元格 `Change` 事件；周期性续约锁。续约失败或 fencing 校验失败时标记会话不可提交，但保留本地文件供用户恢复。
- 提交时先验证会话、锁和结构；取得共享提交互斥后读取 Remote，调用 Diff/Merge。存在冲突时释放短期提交互斥并进入 `ConflictReview`；UI 返回全部选择后重新验证锁/结构并重新读取 Remote，避免裁决期间 Remote 再次变化被覆盖。
- 用户取消裁决时不得写共享工作簿，回到 `CheckedOut`，保留本地会话和列锁。全部冲突解决后才允许声明列写回和 `Save()`；成功确认后关闭句柄、释放列锁并清理本地工作区。
- 放弃前由 UI 确认丢弃；关闭本地副本不保存，释放锁并清理工作区。各清理失败要记录具体未完成项，不得报告为完全成功。

## 3. 涉及函数与模块

模块：`src/MdlSession/`，按 `AGENTS.md` 可协调各底层模块；不直接访问 Excel 对象。

建议过程接口：

```text
Public Function BeginCheckout(targets As Variant, ByRef err As AppErrorInfo) As Boolean
Public Function SubmitChanges(ByRef conflicts As Variant, ByRef err As AppErrorInfo) As SubmitOutcome
Public Function ResumeAfterConflictReview(choices As Variant, ByRef err As AppErrorInfo) As SubmitOutcome
Public Function CancelConflictReview() As Boolean
Public Function AbandonSession(userConfirmed As Boolean, ByRef err As AppErrorInfo) As Boolean
Public Function GetSessionView() As SessionView
Private Function TransitionTo(nextState As SessionState, ByRef err As AppErrorInfo) As Boolean
Private Sub RollbackCheckout(acquiredLeases As Variant, localHandle As String, workspace As String)
```

会话数据至少保存 session ID、状态、目标列清单、Base/结构指纹、锁 lease、工作区路径和本地/远端句柄 ID；不得把 Excel 对象传给 `MdlUI`。

## 4. 功能范围

做：完整生命周期编排、合法状态迁移、列锁续约/校验、签出回滚、冲突取消保留、失败时保留本地数据、成功/放弃清理。

不做：多进程会话恢复 UI、实时协同、自动解决冲突、结构编辑、静默忽略未声明列改动、自动重试不确定结果的保存、全量差异预览。

## 5. 测试方案

- `test/MdlSession/` 先用替身接口测试合法/非法状态迁移、签出各步失败回滚、整组锁失败不留部分会话、续约失败禁提交。
- Excel 集成测试覆盖签出→改值→无冲突提交、冲突进入裁决、取消裁决不写共享且保留本地和锁、裁决完成后仅声明列写回、提交失败保留数据、放弃确认与取消。
- 覆盖重复点击、状态中提交/放弃、结构变化、锁过期、写入/保存/清理失败；Excel 进程和测试均必须受控超时。

## 6. 风险与未决问题

- **冲突裁决期间 Remote 可能继续变化。** 暂定裁决结束后重新读取 Remote，并基于原 Base/最新 Local/新 Remote 再合并；若冲突集合变化则重新展示，不直接套用过期选择。需确认该重复裁决行为的用户体验。
- **保存报错不一定等价于共享文件未改变。** 暂定将其视为结果不确定：保留本地工作区与列锁，禁止盲目重试；重新读取 Remote 并校验后才能继续。需在 Phase 9 确认错误分类与解除方式。
- **本地崩溃残留的恢复/清理规则未定。** 暂按下次启动发现未完成 manifest 时保留文件、记录会话信息并提示人工处理，不自动删除或自动提交；需回来确认是否需要恢复入口和最长保留期。
- 多列锁冲突按已确认的整组取消执行。不同列同时提交通过 Phase 5 的短期提交门闩串行处理，避免共享工作簿同时保存；门闩只覆盖共享访问阶段，不阻止并行本地编辑。
