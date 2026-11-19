# Phase 5：列锁

**依赖：Phase 2。**

## 1. 阶段目的

提供共享目录上的列级锁能力，使同一 sheet + 单个 AMG 列同一时刻只能由一个有效会话编辑，同时允许不同 AMG 列并行编辑。完成后可批量获取、续约、校验和释放锁，并对异常退出的锁执行受控过期回收。

## 2. 实现方式

- 锁目标输入为 N × 3 Variant 矩阵（UNC 文件路径、sheet 名、AMG 表头）。按大小写归一后的完整目标键排序，并以稳定短哈希生成目录名；owner 元数据保存完整目标键，遇到哈希碰撞即失败关闭。owner 文件包含 session ID、GUID fencing token 和版本，过期时间根据共享 owner 文件的最后写入时间计算。
- 暂定列锁 TTL 为 300 秒、续约间隔 60 秒，提交门闩 TTL 为 600 秒；这是需 Phase 9 在目标 Samba 上验证的保守初值。失联、损坏 owner 或时间无法读取均按占用处理。过期回收先将锁目录改名为唯一回收目录，再删除其中的 owner 元数据和空目录。
- 锁协调实现放在 MdlLock 下的 CLockCoordinator 类中；MdlLock 标准模块只暴露产品入口。类仅向同一 VBA 工程提供 Friend 测试入口以注入 LOCALAPPDATA 临时根目录，生产入口不接受本地路径覆盖参数。

- `MdlLock` 管理锁语义，`MdlStorage` 封装目录/文件系统操作，`MdlInfra` 提供错误与日志。调用方提供已规范化的 sheet/AMG 锁目标，锁键包含 UNC 主文件标识、sheet 名和单个 AMG 列名；`MdlLock` 不依赖 `MdlSchema`。
- 用 `MkDir` 对锁目录的创建竞争进行互斥；锁目录中的 `owner.txt` 记录会话标识、所有者、创建/续约时间及 fencing 标识。续约更新固定心跳记录，不创建新的共享临时文件。
- 批量锁按规范化键排序后逐一尝试，避免不同顺序造成互相等待。任一目标失败时，释放本次已取得的锁并整体失败，不保留部分签出；不跳过被占用的列。
- 编辑期列锁之外，提供一个短时共享提交门闩，串行化打开/写入/保存同一共享工作簿的过程；门闩不改变不同列可并行编辑的合同。
- 每次续约、释放和提交前均核对 owner 与 fencing 标识。锁状态不可读、owner 记录损坏或共享访问不确定时按“仍被占用/不可安全接管”处理；只有确认心跳超过配置 TTL 后才尝试回收。
- 所有共享文件系统操作由 `MdlStorage` 执行。共享上只创建锁目录及其必要的 owner/心跳元数据，不创建业务数据临时文件或探测文件。
- `MdlStorage` 的本地工作区/副本接口在本阶段先冻结调用契约，Phase 5 与 Phase 6 可并行实现；联调时两侧必须使用同一契约。

## 3. 涉及函数与模块

模块：`src/MdlLock/MdlLock.bas` 对外提供锁操作，`src/MdlLock/CLockCoordinator.cls` 持有协调器实现（依赖 `MdlInfra`、`MdlStorage`）；`src/MdlStorage/`（依赖 `MdlInfra`）。

当前模块接口：

```text
Public Function AcquireLocks(ByRef targets As Variant, ByVal sessionId As String, ByRef leases As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function RenewLocks(ByRef leases As Variant, ByRef renewed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ValidateLease(ByRef leaseFields As Variant, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function InspectLockStatus(ByRef targetFields As Variant, ByRef status As LockAvailability, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ReleaseLocks(ByRef leases As Variant, ByRef failures As Variant) As Boolean
Public Function TryReclaimExpiredLock(ByRef targetFields As Variant, ByRef reclaimed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function AcquireCommitGate(ByVal fileKey As String, ByVal sessionId As String, ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ValidateCommitGate(ByRef gateFields As Variant, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function ReleaseCommitGate(ByRef gateFields As Variant, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function StorageCreateDirectory(ByVal path As String, ByRef created As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function StoragePathExists(ByVal path As String, ByRef exists As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function StorageReadText(ByVal path As String, ByRef contents As String, ByRef lastWrite As Date, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function StorageWriteText(ByVal path As String, ByVal contents As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function StorageRenameDirectory(ByVal sourcePath As String, ByVal destinationPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function StorageRemoveLockDirectory(ByVal path As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function CreateLocalWorkspace(ByVal sessionId As String, ByRef workspacePath As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function CopySharedFileToWorkspace(ByVal sourceUncPath As String, ByVal workspacePath As String, ByRef localPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
Public Function RemoveLocalWorkspace(ByVal workspacePath As String, ByRef errorInfo As AppErrorInfo) As Boolean
```

`CLockCoordinator` 的内部接口与关键私有过程：

```text
Friend Function InspectTargetLock(ByRef targetFields As Variant, ByRef status As LockAvailability, ByRef errorInfo As AppErrorInfo) As Boolean
Friend Sub SetTestRoot(ByVal path As String)
Friend Sub SetTestTtl(ByVal columnSeconds As Long, ByVal gateSeconds As Long)
Friend Sub SetTestClockOffset(ByVal seconds As Long)
Private Function BuildTargetKey(ByVal fileKey As String, ByVal sheetName As String, ByVal amgHeader As String, ByRef targetKey As String, ByRef lockPath As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function TryAcquireLockDirectory(ByVal lockPath As String, ByVal targetKey As String, ByVal sessionId As String, ByVal ttlSeconds As Long, ByRef fencingToken As String, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ReclaimLockDirectory(ByVal lockPath As String, ByVal expectedKey As String, ByVal ttlSeconds As Long, ByRef reclaimed As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function ValidateLeaseScalars(ByVal lockPath As String, ByVal targetKey As String, ByVal sessionId As String, ByVal fencingToken As String, ByVal ttlSeconds As Long, ByRef valid As Boolean, ByRef errorInfo As AppErrorInfo) As Boolean
Private Function HashKey(ByVal value As String) As String
Private Sub SortTargetOrder(ByRef order() As Long, ByRef targetKeys() As String, ByVal targetCount As Long)
```

`SetTestRoot`、`SetTestTtl`、`SetTestClockOffset` 仅由同工程测试调用，不对加载项外部开放；`MdlLock.InspectLockStatus` 通过 `Friend InspectTargetLock` 获取只读状态，其余锁操作仍经标准模块门面进入。

lease 为 N × 4 Variant 矩阵，列依次为锁路径、完整目标键、session ID 和 fencing token；ValidateLease 接收其中一行的一维数组。commit gate 为一维四项数组（路径、完整目标键、session ID、token）。`InspectLockStatus` 是只读提示，返回空闲、有效占用、过期或未知；损坏 owner/权限错误按未知占用展示，不回收锁。签出时的 `AcquireLocks` 仍是唯一权威判定，状态展示和实际获取之间允许发生竞争。

## 4. 功能范围

做：每个 sheet+单个 AMG 列的锁、批量全有或全无获取、续约、owner 校验、fencing、过期回收、短时提交门闩、`%LOCALAPPDATA%` 工作区创建/复制/清理、失败释放、用本机临时目录模拟文件系统竞争。

不做：表级锁、阻止不同列并行编辑、用户管理面板、强制接管有效锁、共享工作簿内容读写、共享临时文件、依赖映射盘符。

## 5. 测试方案

- `test/MdlLock/` 与 `test/MdlStorage/`：当前组件测试覆盖同键竞争、不同列独立、批量冲突回滚、续约、owner/token 不匹配、过期回收、损坏 owner 保留、提交门闩串行获取和本地工作区约束。权限不足、并发回收竞态、Samba 时间戳/目录原子性与真实双客户端场景留到 Phase 9。
- 使用本地临时目录做组件测试；对 `MkDir` 原子性及实际 Samba 时间戳/删除语义标注为 Phase 9 双客户端验证前提。
- 检查失败路径没有留下本次会话持有的锁；无法证明安全回收时必须拒绝接管。
- 所有锁/存储组件测试在 `Cleanup` 路径删除本次创建的本地临时根目录，并将锁释放/目录删除失败写入测试结果；测试结束后检查 `%LOCALAPPDATA%\MDDTLock\tests` 无本次遗留目录。
- 2026-11-19 本轮验证：`TestMdlStorage.RunMdlStorageTests`、`TestMdlLock.RunMdlLockInspectStatusTest`、`TestMdlLock.RunMdlLockTests` 和完整 `TestSuite.RunAllVbaTests` 均在 Excel 16.0 build 19127 的受控 runner 中通过。runner 对 Excel 退出使用 20 秒硬宽限，完成后未发现遗留 Excel 进程。被超时中止的早期诊断运行曾留下本地锁目录，已通过范围受限清理移除；锁/存储测试的成功和失败路径现统一清理自身临时根目录。Samba 并发语义仍待 Phase 9 外部验证。

## 6. 风险与未决问题

- 多列签出策略已确认：任一目标校验或锁获取失败都整组取消，释放本次已取得锁，不跳过占用列。
- **租约参数是待实测初值，不是生产结论。** 当前暂定列锁 TTL 300 秒、续约间隔 60 秒、提交门闩 TTL 600 秒；以 owner 元数据最后写入时间超过 TTL 判为过期。Phase 9 需在目标 Samba 环境验证时间戳精度、客户端时钟偏差和长时间保存行为后确认或调整。无法可靠读取时间戳时按仍被占用处理。
- **提交门闩的租约与列锁的生命周期边界未定。** 暂定门闩只覆盖共享 Remote 读取、写回和 `Save()`，失败/超时走同一 fencing 校验并释放门闩；需由 Phase 9 真实 SMB 并发测试确认是否足以防止互相覆盖。
- Samba 上 `MkDir` 原子性、目录时间戳精度和删除行为需实机验证；验证失败前不能宣称崩溃锁可自动回收。
- 锁状态检查和本地存储测试的自动化挂起/退出问题尚未定位；需先复现并修复，随后重跑完整锁、存储和聚合测试。
