# 当前实现的用户流程与 VBA 行为

> 本文根据 `src/` 中当前代码整理，核对日期：2026-11-20。它解释现有行为，不替代 `plans/PLAN.md` 和 Phase 计划中的目标合同；发现冲突时应先核实并同步计划或实现。

## 1. 流程概览

```text
手工维护本机配置 -> Start Editing -> 选择工作簿 -> 选择 sheet/AMG 列
-> 预检并批量取锁 -> 复制到 LOCALAPPDATA，记录 Base -> 本地编辑
-> Submit：验证列锁、取得提交门闩、读取 Remote 并三方合并
   -> 无冲突：写声明列并 Save()
   -> 有冲突：逐组件选 Local/Remote，重新读取并核验后写入
-> 关闭本地副本、删除工作区、释放列锁
```

每次会话只操作一份工作簿及用户选择的若干 `sheet + AMG 列`。列锁允许不同列并行编辑；提交期间另有工作簿级短时门闩，串行化共享文件的读取、写入和保存。

## 2. 配置和启动

加载项没有配置管理窗体。每位用户手工维护 `%LOCALAPPDATA%\MDDTLock\config.ini`，自行创建文件夹和配置。版本 2 为每份工作簿写一行；版本 1 的单路径格式仍可读取：

```ini
Version=2
WorkbookPath=\\server\share\team-a\mddt_alpha.xlsx
WorkbookPath=X:\team-b\mddt_beta.xlsx
```

新增工作簿时增加 `WorkbookPath=` 行，移除时删除对应行。加载项不自动创建、补全或修复配置；缺少配置时提示配置路径及创建方式。读取器忽略空行和 `;`、`#` 注释，要求版本为 1 或 2 且至少有一条有效路径，并拒绝格式错误或忽略大小写后的重复路径字符串。

路径语法接受 UNC 和绝对盘符路径，将 `/` 转成 `\` 并去除首尾空格；不探测盘符映射，不解析 UNC/盘符别名。绝对本机路径也会通过语法校验，仅供受控测试；生产配置应指向共享工作簿。不同路径别名能互锁的前提是它们实际到达同一共享目录且具有共同访问权限。

点击 Ribbon **Start Editing** 后，`MdlUI.RibbonStartEdit` 重新读取配置并显示单选工作簿窗体。选定后，`MdlSession.GetAvailableTargets` 确认路径仍在配置清单中，再由 `MdlExcel` 打开共享工作簿以扫描工作表和锁状态。文件/目录不存在、拒绝访问、网络不可达、路径指向文件夹会有分类错误。若 Excel 以只读方式打开则拒绝继续，但提示不能确定是其它 Excel 会话、只读属性还是权限造成。若同一路径已在宿主 Excel 实例打开，也拒绝重复打开。

## 3. 工作表和目标选择

`MdlExcel.ListWorksheetNames` 枚举工作表，`ReadSheetSchema` 读取第 1 行并从右向左定位最后一个非空表头。`MdlSchema.ParseHeader` 将以 `AMG` 开头（不区分大小写）的文本列识别为可编辑列，其它非空文本列视为 product 列；AMG 名称大小写归一后重复会报错。A–E 内的 product 列作为行锚点；没有锚点的 sheet 被跳过并记警告。锚点唯一性是业务前置条件，代码不检测。

每个 AMG 目标显示 sheet、AMG 列名和只读锁状态：`Free`、`Occupied`、`Lease expired; reclaim may be attempted` 或 `Unknown; verified during checkout`。状态仅供提示，读取异常时显示 Unknown，真正结果以取锁时的目录竞争为准。目标窗体允许多选，至少选择一列后点击 **Start Editing**；取消或未选择不启动签出。

## 4. 签出与列锁文件

`MdlSession.BeginCheckout` 再读配置，确认工作簿仍登记、所有目标指向该工作簿，按 sheet 分组，并预检目标 sheet 的结构和可编辑值。之后关闭共享工作簿，再批量取锁。锁按规范化目标键排序；任一目标占用或不能安全回收时，本批已取得的列锁整体回滚，不会只签出空闲部分。

生产锁路径为：

```text
<工作簿目录>\history\mddtlock\<大写并编码的工作簿文件名>\S-<大写并编码的sheet>__C-<大写并编码的AMG>_LOCK\owner.txt
```

选择 `Study` 的 `AMG01` 与 `AMG02` 时会有两个独立锁目录、各自一个 `owner.txt`。工作簿名、sheet、AMG 先转大写；字母、数字、空格、`-`、`_`、`.` 原样保留，其它 UTF-16 字符编码为 `~` 加四位十六进制值。路径中不安全字符不会直接进入目录名。

`CLockCoordinator.BuildLockPath` 按需创建共享端 `history`、`mddtlock`、工作簿名父目录。`TryAcquireLockDirectory` 用 `MkDir` 创建具体 `_LOCK` 目录来仲裁竞争；成功后生成 GUID fencing token 并写入：

```text
Version=1
Session=<session GUID>
Token=<fencing GUID>
Target=<目标键的 UTF-16 十六进制编码>
```

Target 由工作簿文件名、规范化 sheet 名及 AMG 名组成，不含 UNC 或盘符拼写。`owner.txt` 不含用户名，也不含显式创建/续约时间；session/token/target 用于归属校验，文件最后写入时间就是租约心跳。列锁 TTL 为 300 秒，每 60 秒续约，续约时重写同一 `owner.txt`。

目录已存在时，程序读取 owner 并核对目标键及最后写入时间。有效且未过期视为占用；owner 缺失、损坏、不可读或目标不匹配时按占用处理。确认超过 TTL 后，先把具体锁目录改名为 `reclaim-<GUID>`，再删除 owner 和空目录，随后重新创建锁目录。查看状态不会回收，只有实际取锁时尝试回收。

批量失败回滚或正常释放只删除具体锁目录及其中的 owner，保留父目录，没有全树清理。由于父目录在竞争前创建，签出失败也可能留下空父目录；回收重命名成功但删除失败则可能留下 `reclaim-<GUID>` 目录。共享端不创建业务副本或提交临时文件。

## 5. 本地副本、快照和编辑

取得全部列锁后，`MdlStorage.CreateLocalWorkspace` 在 `%LOCALAPPDATA%\MDDTLock\sessions\<session GUID>\` 建立工作区，`CopyWorkbookToWorkspace` 将整个共享工作簿复制到该目录并保留文件名；`MdlExcel.OpenWorkbookHandle` 在宿主 Excel 实例打开副本。`MdlSession` 为每个目标 sheet 读取 Base 快照后进入 `CheckedOut`。复制、打开或快照失败会关闭本地句柄、删除工作区并回滚本批锁。

快照从第 2 行开始，末行是 A–E 锚点列与所选 AMG 列各自最后非空行的最大值，内部空行保留物理行位置。Base 保存表头指纹、边界、列索引、锚点位置和值，以及声明范围的值、字体和填充数据。边界不使用 `UsedRange` 或 `Range.Find`。AMG 区域不允许公式，只接受字符串或真正空白；`Empty` 和零长度字符串比较不同。锚点不写回，但提交前用来检查行映射是否变化。

用户在本地副本中编辑声明的 AMG 列。编辑期间不访问共享主文件，也不依靠 `Change` 事件。会话处于 CheckedOut 或 ConflictReview 时，`Application.OnTime` 每 60 秒调用 `SessionHeartbeat`：核验 session/token/target/TTL，然后重写同一 owner 文件更新时间戳。续约失败会禁用 Submit 并记录日志，不自动重新取锁。异常退出不会正常释放锁；超过 TTL 后，由后续签出尝试回收。

## 6. 提交、合并与冲突复核

点击 **Submit** 后，`MdlSession.ExecuteCommit` 先核验全部列锁，再获取工作簿级 `S-COMMIT__C-GATE_LOCK\owner.txt` 门闩（TTL 600 秒），随后打开共享 Remote。门闩只串行化提交阶段，不影响各列本地并行编辑。程序分别对 Local 和 Remote 检查表头/结构、数据行边界、锚点列和锚点值；任一结构或锚点映射变化即停止写入。

提交读取 Base、当前 Local、最新 Remote。`MdlDiff.CompareSnapshots` 比较 Local 与 Base；`MdlMerge.MergeThreeWay` 只对声明列逐行处理，先以 Remote 初始化结果，再按组件应用三方规则：Local=Base 取 Remote；Remote=Base 取 Local；Local=Remote 取共同值；三者都不同则冲突。

| 组件 | 比较范围 | 冲突处理 |
|---|---|---|
| `TextFont` | 字符串值及绑定到字符位置的直接字体 run；空白/零长度文本使用单元格级字体 | 文本和值与字体 run 一起作为整体，不能把两边不同字符的格式拼接 |
| `Fill` | 单元格级直接填充 | 独立于 TextFont 合并和裁决 |

因此同一单元格可能产生 TextFont 与 Fill 两条裁决记录。值按字符串精确比较，大小写/空格差异算变化，Empty 与 `""` 不等。边框、对齐、数字格式、条件格式等不参与。

首次发现冲突时不写共享文件；进入 ConflictReview，关闭 Remote 并释放提交门闩，保留本地副本、Base 和列锁。冲突窗体逐条展示 sheet/单元格/组件及 Base、Local、Remote 预览；用户必须为每条选 **Keep Local** 或 **Keep Remote**，全部选择后才可应用，可前后浏览。Cancel 或关闭窗体回到 CheckedOut，不写共享文件，保留本地改动和列锁。

选择完成后，程序重新获取门闩、重开 Remote 并重算。如果复核期间冲突视图变化，丢弃旧选择并显示新冲突要求重新裁决。视图一致时，`MdlMerge.ApplyConflictChoices` 验证每个冲突恰好被选一次且选择合法。写入前再次核验列锁和门闩；`MdlExcel.WriteDeclaredColumns` 只写声明列的合并值、逐字符字体和填充，并回读值核验。写入后、保存前再次校验锁，最后对共享工作簿调用 `Workbook.Save()`，不调用 `SaveAs`。

若 Local 相对 Base 没有变化，即使 Remote 有更新，也不写共享文件或保存，直接结束本次会话并清理资源。

## 7. 完成、失败与保留行为

| 情况 | 共享文件 | 本地副本和列锁 | 状态 |
|---|---|---|---|
| 无本地变化或保存成功 | 无变化或已保存合并结果 | 关闭并清理本地工作区，核验 owner 后释放具体列锁 | Idle；清理不完整时提示共享提交已完成但资源清理有警告 |
| 保存前校验、打开、读取、合并或写入失败 | 尚未调用 Save；关闭 Remote 时不保存 | 保留本地副本和列锁，释放提交门闩 | 通常回到 CheckedOut，可在状态允许时重试 |
| 取消冲突复核 | 不写共享文件 | 保留本地副本和列锁 | 回到 CheckedOut，可稍后重试 |
| Save 失败/异常 | 结果可能已写或未写，无法从错误断定 | 保留本地副本和列锁，禁用再次 Submit，尝试释放门闩 | 要求先核查共享文件，不自动重试；确认后可 Abandon |
| 确认 Abandon | 不提交本地修改 | 不保存地关闭副本、删除工作区、尝试释放列锁 | Idle；无法释放的锁待过期回收 |
| Excel/进程意外退出 | 无正常清理保证 | 本机目录可能残留，owner 保留到回收 | 心跳停止，TTL 到期后下次获取锁尝试回收 |

释放锁仍要匹配 owner/token 且租约未过期；过期或所有权变化时不强行删除。工作区清理失败可能遗留本机文件，锁释放失败的目录由 TTL 和之后的取锁回收处理。

## 8. 计划与实测核对

- 总计划已概述用户流程及系统响应；本文补充过程顺序、具体路径、锁元数据、合并组件和失败副作用。
- Phase 5 曾称 owner 文件含“所有者、创建/续约时间”。源码实际字段只有 `Version`、`Session`、`Token`、`Target`；没有用户名/显式时间，时间来自文件最后写入时间。该段计划已纠正。
- Ribbon、窗体和消息字符串当前为英文，如 `Start Editing`、`Submit`、`Keep Local`。`AGENTS.md` §4.1 规定面向用户字符串中文，与源码不一致。本次只记录现状，没有擅自翻译 UI 或修改规范。
- 已记录的本机证据包括真实临时 `.xlsx` 导入、声明列修改、Save 回读、有效/损坏假锁拦截，以及 Ribbon 和工作簿选择窗体显示/取消冒烟。完整目标选择、冲突窗体交互、端到端签出到提交、代表性业务模板、真实 Samba 双客户端锁竞争和共享保存仍未验收。

主要入口：`MdlUI`、`MdlSession`、`CLockCoordinator`、`MdlStorage`、`MdlExcel`、`MdlSchema`、`MdlDiff`、`MdlMerge`，分别位于 `src/` 下同名目录。
