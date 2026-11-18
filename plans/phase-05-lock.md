# Phase 5：列锁

**依赖：Phase 2。**

## 1. 阶段目的

提供共享目录上的列级锁能力，使同一 sheet + 单个 AMG 列同一时刻只能由一个有效会话编辑，同时允许不同 AMG 列并行编辑。完成后可批量获取、续约、校验和释放锁，并对异常退出的锁执行受控过期回收。

## 2. 实现方式

- `MdlLock` 管理锁语义，`MdlStorage` 封装目录/文件系统操作，`MdlInfra` 提供错误与日志。调用方提供已规范化的 sheet/AMG 锁目标，锁键包含 UNC 主文件标识、sheet 名和单个 AMG 列名；`MdlLock` 不依赖 `MdlSchema`。
- 用 `MkDir` 对锁目录的创建竞争进行互斥；锁目录中的 `owner.txt` 记录会话标识、所有者、创建/续约时间及 fencing 标识。续约更新固定心跳记录，不创建新的共享临时文件。
- 批量锁按规范化键排序后逐一尝试，避免不同顺序造成互相等待。任一目标失败时，释放本次已取得的锁并整体失败，不保留部分签出；不跳过被占用的列。
- 编辑期列锁之外，提供一个短时共享提交门闩，串行化打开/写入/保存同一共享工作簿的过程；门闩不改变不同列可并行编辑的合同。
- 每次续约、释放和提交前均核对 owner 与 fencing 标识。锁状态不可读、owner 记录损坏或共享访问不确定时按“仍被占用/不可安全接管”处理；只有确认心跳超过配置 TTL 后才尝试回收。
- 所有共享文件系统操作由 `MdlStorage` 执行。共享上只创建锁目录及其必要的 owner/心跳元数据，不创建业务数据临时文件或探测文件。
- `MdlStorage` 的本地工作区/副本接口在本阶段先冻结调用契约，Phase 5 与 Phase 6 可并行实现；联调时两侧必须使用同一契约。

## 3. 涉及函数与模块

模块：`src/MdlLock/`（依赖 `MdlInfra`、`MdlStorage`）；`src/MdlStorage/`（依赖 `MdlInfra`）。

建议过程接口：

```text
Public Function AcquireLocks(targets As Variant, sessionId As String, ByRef leases As Variant, ByRef err As AppErrorInfo) As Boolean
Public Function RenewLocks(leases As Variant, ByRef renewed As Boolean, ByRef err As AppErrorInfo) As Boolean
Public Function ValidateLease(lease As LockLease, ByRef valid As Boolean, ByRef err As AppErrorInfo) As Boolean
Public Function ReleaseLocks(leases As Variant, ByRef failures As Variant) As Boolean
Public Function TryReclaimExpiredLock(target As LockTarget, ByRef reclaimed As Boolean, ByRef err As AppErrorInfo) As Boolean
Public Function AcquireCommitGate(fileKey As String, sessionId As String, ByRef gate As CommitLease, ByRef err As AppErrorInfo) As Boolean
Public Function ReleaseCommitGate(gate As CommitLease, ByRef err As AppErrorInfo) As Boolean
Public Function StorageCreateDirectory(path As String, ByRef created As Boolean, ByRef err As AppErrorInfo) As Boolean
Public Function StorageReadOwner(path As String, ByRef owner As LockOwner, ByRef err As AppErrorInfo) As Boolean
Public Function StorageWriteHeartbeat(path As String, timestamp As Date, ByRef err As AppErrorInfo) As Boolean
Public Function CreateLocalWorkspace(sessionId As String, ByRef workspacePath As String, ByRef err As AppErrorInfo) As Boolean
Public Function CopySharedFileToWorkspace(sourceUncPath As String, workspacePath As String, ByRef localPath As String, ByRef err As AppErrorInfo) As Boolean
Public Function RemoveLocalWorkspace(workspacePath As String, ByRef err As AppErrorInfo) As Boolean
```

`LockLease` 向会话返回目标键、owner/session 标识和不可由调用方构造的 fencing 标识；上层不直接拼接锁目录结构。

## 4. 功能范围

做：每个 sheet+单个 AMG 列的锁、批量全有或全无获取、续约、owner 校验、fencing、过期回收、短时提交门闩、`%LOCALAPPDATA%` 工作区创建/复制/清理、失败释放、用本机临时目录模拟文件系统竞争。

不做：表级锁、阻止不同列并行编辑、用户管理面板、强制接管有效锁、共享工作簿内容读写、共享临时文件、依赖映射盘符。

## 5. 测试方案

- `test/MdlLock/` 与 `test/MdlStorage/`：同键竞争仅一方成功、不同列均可成功、批量中途冲突后回滚已取得锁、重复续约/释放、owner 不匹配、过期与未过期、损坏 owner、权限不足、目录删除失败、并发回收竞争、提交门闩串行获取、工作区路径位于 `%LOCALAPPDATA%`、复制/清理失败。
- 使用本地临时目录做组件测试；对 `MkDir` 原子性及实际 Samba 时间戳/删除语义标注为 Phase 9 双客户端验证前提。
- 检查失败路径没有留下本次会话持有的锁；无法证明安全回收时必须拒绝接管。

## 6. 风险与未决问题

- 多列签出策略已确认：任一目标校验或锁获取失败都整组取消，释放本次已取得锁，不跳过占用列。
- **TTL、续约间隔与过期判据未在总计划给定。** 暂以可配置常量表达，过期时间须大于续约间隔并容忍短时 SMB 延迟；具体数值需在 Phase 9 真实环境验收前确认。
- **提交门闩的租约与列锁的生命周期边界未定。** 暂定门闩只覆盖共享 Remote 读取、写回和 `Save()`，失败/超时走同一 fencing 校验并释放门闩；需由 Phase 9 真实 SMB 并发测试确认是否足以防止互相覆盖。
- Samba 上 `MkDir` 原子性、目录时间戳精度和删除行为需实机验证；验证失败前不能宣称崩溃锁可自动回收。
