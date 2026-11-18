# 具体过程：用户视角 / 程序视角

配套文档：`DESIGN.md`（设计定稿）。本文回答"实际用起来是什么样"和"代码里到底发生什么"。

---

## 0. 一句话

用户看到的是：**勾选 sheet → 本地随便改（几分钟到几小时）→ 点提交 → 看预览 → 确认 → 等 10~20 秒**。
服务器文件在这 10~20 秒之外**全程无人占用**，读者随时可开。

---

## 1. 角色与文件

| 角色 | 行为 | 是否占用服务器文件 |
|---|---|---|
| **编辑者** | 声明 sheet → 本地编辑 → 提交 | 只在提交的 10~20 秒内 |
| **只读读者** | 通过"只读查看"按钮或直接打开 | ⚠️ 直接双击打开会占用（见 §8.2） |
| **旁观者** | 打开"谁在改什么"面板 | 不占用 |

| 文件 | 位置 | 用途 |
|---|---|---|
| `book.xlsx` | `\\server\share\` | 唯一数据文件，无宏 |
| `sync.xlam` | `%APPDATA%\Microsoft\AddIns\` | 工具本体，每用户装一次 |
| `owners.json` | `\\server\share\_sync\` | 谁声明了哪张 sheet |
| `.lock` | `\\server\share\_sync\` | 两个互斥锁（声明锁、提交锁） |
| `audit\<user>.log` | `\\server\share\_sync\audit\` | 每人一个文件，避免并发追加踩踏 |
| 本地副本 | `%LOCALAPPDATA%\ExcelSync\book.xlsx\book.xlsx` | 编辑载体 |
| 用户配置 | `%APPDATA%\ExcelSync\config.json` | 文件路径、上次勾选的 sheet |

---

## 2. 状态机

单个 sheet 的状态（`owners.json` 里每个 claim 一条）：

```
   空闲 ──(声明)──> 被A占用 ──(A提交成功)──> 空闲
     ▲                 │
     │                 ├──(A放弃)──────────> 空闲
     │                 ├──(A崩溃/TTL过期)──> 可接管(带警告)
     └──(A强制释放)────┘
```

每个用户会话的状态：

```
   无副本 ──(声明+下载)──> 编辑中 ──(提交成功)──> 无副本 / 再次编辑中
                             │
                             ├──(放弃)──> 无副本（删本地）
                             └──(崩溃)──> 编辑中(残留) ──(下次启动可恢复)
```

---

## 3. 用户视角

### 3.1 一次性配置（每人一次，2 分钟）

1. 双击 `install.cmd`（或手工）：把 `sync.xlam` 拷到 `%APPDATA%\Microsoft\AddIns\`，注册加载项
2. IT 侧（一次性）：把 `\\server\share` 加入 Excel **受信任位置**（关键，否则打开服务器文件会进"受保护视图"，VBA 无法读写）
3. 首次使用点"设置"，选 `\\server\share\book.xlsx`，保存

之后每次开 Excel，功能区多一个 **「协同」** 选项卡：

```
[开始编辑] [只读查看] [提交] [放弃] [谁在改什么] [设置]
```

### 3.2 开始编辑

1. 点 **「开始编辑」**
2. 弹窗 A：**勾选要维护的 sheet**（列表带状态列）

```
┌ 选择要维护的 Sheet ─────────────────────────┐
│  ☐  封面/TOC        空闲                     │
│  ☑  客户主数据       空闲                     │
│  ☐  2024Q1          张三 10:23 起修改中      │
│  ☑  2024Q2          空闲                     │
│  ☐  对账表           (上次勾选) 空闲          │
│  ... 30 张                                   │
│                              [取消] [确定]   │
└──────────────────────────────────────────────┘
```

   - 默认勾选上次的选择；别人占用的自动灰掉

3. 程序抢锁写入声明。**若在勾选期间被别人抢先**：
   > `2024Q1` 已被 张三 于 10:23 声明。是否只保留其余 sheet 继续？（[是] / [取消]）

4. 进度条：`正在下载 book.xlsx ...`（1~2 秒）
5. 本地副本自动打开，标题栏显示 **[协同] book.xlsx - Sheet3, Sheet7**
6. 功能区「提交」「放弃」变可用

**此刻起，用户面对的是一份普通本地 Excel 文件。断网也不影响编辑。**

### 3.3 编辑期间

- 零网络访问，保存到本地，随时 Ctrl+S
- 程序不做任何监视、不加任何事件钩子 → **不卡、不弹窗、不会崩掉用户的编辑**
- 想确认别人在改什么：点「谁在改什么」，只读面板
- **不要求提交前必须保存**（程序提交时会以内存中的当前状态为准）

### 3.4 提交（核心）

点 **「提交」**：

**第 1 步 —— 自检（<0.5s，纯本地）**

对每个声明的 sheet 检查：key 列非空、key 唯一、key 列是文本格式。失败则：

```
┌ 提交前检查未通过 ───────────────────────┐
│ Sheet3  第 42 行 key 为空                │
│ Sheet3  第 87 行 key 与第 12 行重复(K010)│
│ Sheet7  第 3 行 key 为数字格式，建议改文本│
│                          [仍然提交] [返回]│
└──────────────────────────────────────────┘
```

**第 2 步 —— 我的改动预览（<1s，纯本地）**

```
┌ 我的改动 ───────────────────────────────────────────┐
│ Sheet3    新增 2 行   修改 17 行   删除 0 行         │
│ Sheet7    新增 0 行   修改 1 行    删除 0 行         │
│                                                     │
│ Sheet3                                              │
│  行  K0100  [新增]                                  │
│  行  K0102  单元格 D18:  "待定" → "已确认"           │
│  行  K0102  样式   E18:  填充 #FFFFFF → #FFF2CC      │
│  ... 共 19 处  (显示前 200 条)                       │
│                                       [取消] [提交] │
└─────────────────────────────────────────────────────┘
```

   - 行身份按 key 对齐 → 用户插一行不会变成几百条假改动
   - 样式只报：**填充色、字体色**。列宽/行高/边框不报

**第 3 步 —— 打开服务器文件 + 冲突检查（5~10s）**

   - 进度显示：`正在校验服务器文件 ...`
   - 快路径（服务器 mtime+size 与下载时一致）→ 跳过冲突段，直奔第 4 步
   - 否则比对声明 sheet 的区域哈希：
     - 全部一致 → 直奔第 4 步
     - 有不一致 → 追加弹窗（仅此时出现）：

```
┌ 检测到他人改动 ─────────────────────────────────┐
│ Sheet3 在你编辑期间被 李四 修改过                │
│                                                 │
│ · 服务器新增 3 行（K0110/K0111/K0112）           │
│   → 本次提交会覆盖（丢失）这 3 行                │
│ · K0102 双方都改了 → 真冲突                      │
│   服务器: "已确认"   你的: "待商榷"              │
│                                                 │
│ 建议：取消提交，重新拉取后重做                    │
│              [取消提交] [导出我的改动] [强制覆盖] │
└─────────────────────────────────────────────────┘
```

**第 4 步 —— 写入（5~10s）**

   - `正在写入服务器 ...`
   - 成功：

```
  ✅ 提交成功
  Sheet3: 新增 2 行 / 修改 17 行
  Sheet7: 修改 1 行
  用时 14 秒
  本地副本：[关闭] [继续编辑]
```

   - `继续编辑` → 重新拉取最新版本、重新声明同一批 sheet，继续
   - `关闭` → 删本地副本、释放声明

**失败时（网络断、被读者锁住、写盘失败）**：

```
  ⚠ 提交失败：文件被其他用户打开，请稍后重试
  你的改动已保存在本地，不会丢失。
  位置：%LOCALAPPDATA%\ExcelSync\book.xlsx\
  可点「重试提交」或「导出我的改动」另存一份。
```

   > **关键保证：任何失败都不写服务器。** 内存中改动 + 本地副本都在。

### 3.5 放弃 / 崩溃恢复

- 「放弃」→ 二次确认 → 删除本地副本 → 释放声明
- Excel 崩溃后重开：检测到"你有未提交的声明 + 本地副本有新改动" →

```
  检测到上次未提交的编辑（Sheet3，11-15 10:23）
  [继续编辑并提交] [导出到桌面] [丢弃]
```

### 3.6 只读读者

点 **「只读查看」** → 下载到本地 → 以只读方式打开，**不声明、不占用**。
（为什么不直接双击打开：见 §8.2）

---

## 4. 程序视角

### 4.1 模块划分

| 模块 | 职责 |
|---|---|
| `modConfig` | 路径常量、用户配置读写、受信任位置校验 |
| `modLock` | SMB 字节锁封装（`Acquire`/`Release`/`WithRetry`） |
| `modClaim` | `owners.json` 的读-改-写（全程持声明锁） |
| `modWorkspace` | 本地副本下载、base 快照、崩溃恢复 |
| `modKey` | `normalizeKey`、key 提取与自检 |
| `modHash` | 区域哈希（FNV-1a 手写，无外部依赖） |
| `modDiff` | base↔local 的行级/单元格级差异（用于预览） |
| `modGuard` | base↔server 的区域哈希与 key 级分类（用于冲突） |
| `modCommit` | 提交事务：打开→清区→粘贴→校验→保存 |
| `modAudit` | 追加审计日志 |
| `frmPickSheets` / `frmDiffPreview` / `frmConflict` / `frmWhoIsEditing` | 四个窗体 |

### 4.2 「开始编辑」调用序列

```
Ribbon: btnStartEdit
 └─ modConfig.Load()
 └─ modClaim.WithClaimLock()               ' SMB 锁，毫秒级
 │   ├─ ReadOwnersJson()
 │   ├─ frmPickSheets.Show(currentOwners)  ' 窗体显示期间仍持锁？→ 否，先读后放
 │   └─ WriteOwnersJsonAtomic()            ' 写临时 + 覆盖
 └─ modWorkspace.EnsureLocalCopy()
 │   ├─ If FileLen(UNC) == base.size And FileDateTime(UNC) == base.mtime → skip
 │   └─ Copy-Item UNC → %LOCALAPPDATA%\ExcelSync\book.xlsx\book.xlsx
 └─ modHash.RegionHash(local, sheet)       ' 每张声明 sheet，写入 owners.json 的 baseHash
 └─ Workbooks.Open(localPath)
 └─ modConfig.RecordSession(activeSheets, baseStamp, baseHashes)
 └─ Ribbon.Invalidate                       ' 提交/放弃 变可用
```

**声明锁的持有窗口 = 毫秒级**，绝不在窗体显示期间持锁（否则用户看着弹窗发呆会阻塞所有人）。做法：先读 → 放锁 → 显示窗体 → 再抢锁 → 重读 → 校验冲突 → 写入 → 放锁。中间被别人抢先就在第二步发现（用户已被告知）。

### 4.3 「提交」调用序列（核心事务）

```
Ribbon: btnCommit
│
├─【阶段 0：本地自检】modKey.SelfCheck(activeSheets)          ~0.1s
│    失败 → 弹报告，可强制继续
│
├─【阶段 1：我的改动预览】modDiff.Build(local, baseSnapshot)   ~0.5s
│    frmDiffPreview.Show                                      (用户可取消 ★回滚点1)
│
├─【阶段 2：抢提交锁】modLock.Acquire("commit.lock", ttl:=30s)
│    失败重试 20s（每 2s 一次，最多 10 次）
│    超时 → "有人正在提交，请稍后" → 退出（本地无损）
│
├─【阶段 3：打开服务器文件】Workbooks.Open(UNC, ReadOnly:=False)
│    被读者锁住 → 重试 3 次×5s → 失败则报"文件被占用"（★回滚点2：不 Save 直接 Close）
│
├─【阶段 4：冲突检查】modGuard.Check(server, baseSnapshot, activeSheets)
│    ├─ stamp(server) == baseStamp → 跳过（快路径）
│    └─ 否则逐 sheet 算 modHash.RegionHash(server, sheet)
│         ├─ 全部 == baseHash → 通过
│         └─ 有差异 → modGuard.ClassifyByKey() 按 key 分类
│              frmConflict.Show → 用户选 [取消] 或 [强制覆盖]  (★回滚点3)
│
├─【阶段 5：写入】modCommit.Patch(server, local, activeSheets)
│    对每张 sheet：
│      1. 服务器区域 ClearContents + ClearFormats          ' 解决"缩小"残留
│      2. local.Range(localUsed).Copy Destination:=server.Range(A1)
│      3. modCommit.Verify():
│            hash(serverRegion.Value2) == hash(localRegion.Value2)   ' 必须相等
│            hyperlinkCount(server) == hyperlinkCount(local)          ' TOC 保护
│      4. 校验失败 → 立即 Close SaveChanges:=False → 报错退出 (★回滚点4)
│
├─【阶段 6：保存】server.Save                             5~10s
│    失败 → 重试 2 次 → Close SaveChanges:=False → 报错 (★回滚点5)
│
├─【阶段 7：收尾】
│    server.Close
│    modLock.Release("commit.lock")
│    modClaim.Release(activeSheets)          ' 持声明锁，毫秒级
│    modAudit.Append(user, sheets, base, new, mode, rowsAdded/Modified/Deleted)
│    清理本地副本（或按用户选择保留并重新拉取）
```

### 4.4 事务性：为什么这个提交是安全的

> **Excel 的 `Close SaveChanges:=False` 就是回滚。**

从阶段 3（打开服务器文件）到阶段 6（Save）之间，有 5 个 ★ 回滚点。任何一个检查不通过，都执行 `Close SaveChanges:=False`，服务器文件**字节级不变**。

所以提交是**全有或全无**，不存在"写了一半"的中间态。这一点由 Excel 自身的 `~$` + 临时文件 + 重命名机制兜底，比自己 copy 覆盖强得多（后者会制造"别人读到半个文件"的窗口）。

### 4.5 为什么要 Clear 再 Copy（"缩小"陷阱）

用户删了 20 行 → 本地区域 480 行，服务器区域仍 500 行。
若直接 Copy 覆盖 A1:480，**服务器第 481~500 行会残留**，看起来像"删除没生效"。

所以必须先 `ClearContents + ClearFormats` 服务器区域，再粘贴。清理范围取 `Union(base 区域, server 当前区域)` 的并集。

**已实测确认**（`VERIFIED.md §3`）：不清除时 A5/A10 明确残留旧值；清除并集后完全干净。

另外两条实测补充：
- `Range.Copy` **不携带行高/列宽**（Excel 固有行为），需 Copy 后显式赋值（3 行代码）。你说列宽不关心，顺手加上即可。
- **不要用 VBA 的 `Name` 语句做重命名覆盖** —— 目标已存在时报 `error 58`。若要"重命名换页"，必须用 `MoveFileExW(MOVEFILE_REPLACE_EXISTING)` 或 FSO `MoveFile(OverWriteFiles:=True)`。

---

## 5. 时间线（实测预期，10MB / 30 sheet）

| 步骤 | 耗时 | 说明 |
|---|---|---|
| 点「开始编辑」→ 勾选完成 | 人决定 | — |
| 写声明 | <0.1s | 毫秒级 SMB 锁 |
| 下载 10MB | **0.5~2s** | 机房内网，不含 Excel 解析 |
| 本地打开 | 5~10s | Excel 自己的解析时间，非网络 |
| **编辑** | 分钟~小时 | 零网络 |
| 自检 + 我的改动预览 | 0.5~1s | 纯本地 |
| 打开服务器文件 | 5~10s | Excel 解析，**占用开始** |
| 冲突检查 | <1s | 只算声明的几张表 |
| Clear + Copy + 校验 | 0.5~2s | 约 6,250 格 |
| Save | 5~10s | — |
| **服务器占用总时长** | **约 10~20s** | 原方案是小时级 |

> 已实测（`VERIFIED.md`）：`Range.Copy` 36 格耗时 **62ms**；500×50=25,000 格耗时 **62ms**。
> 你的规模约 6,250 格 → Copy 环节 <100ms，可忽略。瓶颈只在 Excel 打开/保存。
>
> 若做文件瘦身（10MB → 2MB），打开/保存各降到 1~3s，占用窗口可压到 ~5s。
> 这一步不阻塞，但收益是双份的（用户自己开文件也更快）。

---

## 6. 锁的具体实现（SMB 字节锁，不是"写个文件猜")

```vba
' 真实原子互斥：Windows 映射盘 → SMB2 byte-range lock
Function Acquire(lockPath As String, ttlSeconds As Long) As Boolean
    On Error Resume Next
    Close #1                              ' 确保干净
    Open lockPath For Binary Access Write Lock Read Write As #1
    If Err.Number = 0 Then Acquire = True: Exit Function
    ' Err 70 = Permission denied → 别人持有
    Err.Clear
    If IsStaleLock(lockPath, ttlSeconds) Then
        Kill lockPath                     ' 仅在 TTL 超时后才敢这么做
        Open lockPath For Binary Access Write Lock Read Write As #1
        Acquire = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function
```

要点：

- `Lock Read Write` 由**服务器端**强制 → 跨进程、跨机器真互斥，不是"检查文件存在"式的竞态
- 客户端崩溃时 SMB 会话断开，服务器自动释放字节锁 → 比纯 TTL 文件更可靠
- TTL 只用于兜底"句柄泄漏"这种罕见情况（同一 RDP 主机上进程僵死）
- **锁文件里写内容**（用户/时间/PID）仅供人阅读，不承担互斥语义

`owners.json` 的所有读写都必须走这个锁；写入用"临时文件 + 覆盖"，因为全程持锁，不需要真正的原子重命名。

---

## 7. 哈希实现（无外部依赖）

VBA 没有内置 SHA。**不要**用 `CreateObject("System.Security.Cryptography.SHA256Managed")`（依赖 mscorlib 注册，部分环境被封）。手写 FNV-1a 64 位，约 20 行：

```vba
' 区域哈希：只含值，不含样式，不含列宽行高
Function RegionHash(ws As Worksheet) As String
    Dim lastRow&, lastCol&, v, buf$, r&, c&
    lastRow = LastUsedRow(ws): lastCol = LastUsedCol(ws)
    If lastRow = 0 Then RegionHash = "empty": Exit Function
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2
    buf = lastRow & "x" & lastCol & "|"
    For r = 1 To lastRow
        For c = 1 To lastCol
            buf = buf & NormalizeCell(v(r, c)) & Chr(30)
        Next c
    Next r
    RegionHash = Fnv1a64(buf)
End Function
```

- 非对抗场景，FNV-1a 碰撞概率可忽略；我们要的是"变没变"，不是密码学
- 只哈希**声明的 sheet**（1~5 张），不是 30 张 → 常驻开销 <1s
- 序列化时必须用分隔符（`Chr(30)`）并区分"空"与"空字符串"，否则 `A1="a",A2="bc"` 与 `A1="ab",A2="c"` 会撞哈希

---

## 8. 三个必须修正/新增的点（对 DESIGN.md 的修订）

### 8.1 L1"快路径"不能作为正确性守卫

原设计里 `mtime + size` 相同就直接覆盖 —— **这个不能保证内容没变**（同一秒内改回原大小、或 `mtime` 被工具重置）。

修正：

- **正确性守卫只有区域哈希**（阶段 4）
- `mtime + size` 降级为**UI 排序优化**：用它判断"是否需要显示冲突段"，从而让"我的改动预览"能排在打开服务器文件之前（用户少等 10 秒）

### 8.2 读者直接双击打开，会阻塞提交（最大的运营风险）

**这是当前设计最脆弱的一环，必须正面处理。**

Excel 默认以**读写方式**打开文件并持有锁。任何人在编辑时段双击 `\\server\share\book.xlsx` 只是"看一眼"，也会锁住整个文件 → 别人提交时 `Workbook.Open` 失败。

"服务器文件只在提交时被占 10~20 秒"这个承诺，**要求所有人（包括纯读者）都不以读写方式直接打开**。这是协议依赖，不是技术保证。

缓解措施（建议全做）：

1. 工具提供 **「只读查看」** 按钮（下载本地 → 只读打开，不占用）
2. 分发一个指向"只读查看"的 `.lnk` 或说明文档，替代"直接双击"
3. 提交失败时提示**具体是谁**在占用（`~$` 文件名 = 用户名），让人去找他
4. 提交侧自动重试（3 次 × 5s），覆盖"短暂打开又关掉"的情形

### 8.3 可选的架构升级：写临时 + 重命名换页（需先实测）

若你的 Samba 允许"覆盖一个正被其他客户端打开的文件"（POSIX `rename()` 语义；受 `strict rename` 影响，默认 `no` 通常允许），则可以做到**提交时完全不锁 book.xlsx**：

```
1. 下载 book.xlsx → 本地临时
2. 本地打补丁 → 本地另存
3. 上传到 _sync\upload\<uuid>.xlsx
4. MoveFileReplace(upload, book.xlsx)     ' 读者手里的旧 inode 不受影响
```

好处：读者爱开多久开多久，永不阻塞提交；提交窗口从 20s 降到 ~2s。
代价：读者若**保存**会把整个提交冲掉（所以读者必须只读，协议要求不变，但不再有"锁"问题）。

**先做 §10 的实验 3 再决定是否采用。**

---

## 9. 失败路径总表

| 场景 | 检测点 | 用户看到 | 服务器 |
|---|---|---|---|
| 勾选时被抢先声明 | 阶段 3.2 二次抢锁 | "已被 X 声明" | 不变 |
| 下载失败 | `Copy-Item` 异常 | "网络异常，请重试" | 不变 |
| 提交时别人正在提交 | 抢提交锁超时 20s | "有人正在提交" | 不变 |
| 提交时读者开着文件 | `Workbook.Open` 失败 | "文件被 X 占用" | 不变 |
| 他人改了我的 sheet | 阶段 4 哈希不符 | 冲突弹窗，默认取消 | 不变 |
| key 不合格 | 阶段 0 自检 | 检查报告 | 不变 |
| 粘贴后校验不一致 | 阶段 5.3 | "内部校验失败，已放弃" | **不变**（`Close False`） |
| Save 失败 | 阶段 6 | "写入失败，改动已留本地" | 不变或完整 |
| Excel 崩溃 | 下次启动 | "检测到未提交的编辑" | 不变 |
| 声明残留（TTL 过期） | 「谁在改什么」 | "可强制接管" | 不变 |

---

## 10. 需要先跑的 3 个实验（各 10 分钟，决定实现细节）

| # | 验证什么 | 怎么测 | 影响 |
|---|---|---|---|
| 1 | `Range.Copy Destination:=` 跨工作簿是否完整保留**超链接 + 填充色 + 字体色** | 手工：两本工作簿，TOC 带跳转链接，Copy 过去后检查 `Hyperlinks.Count` 和 `Interior.Color` | 若不保留 → 改为逐属性写入 + 显式重建超链接 |
| 2 | 服务器文件被 A 打开时，B 能否 `Open ReadOnly:=False` | 两人同时开 | 决定重试策略文案 |
| 3 | A 开着 book.xlsx 时，B 能否把另一个文件 `Name`/`MoveFile` 覆盖到 book.xlsx | 命令行 `move /y` 测试 | 通过 → 采用 §8.3 架构（体验大幅提升） |

实验 1 是**技术前提**，实验 3 是**架构升级的钥匙**，都值得先做。

---

## 11. 还未定的两项（不阻塞本文）

- Excel 版本与位数 → 影响 `Range.Copy` 与 COM 细节，但三个实验会顺带暴露
- TOC 是否有按钮/定义名称/冻结窗格 → 若有，TOC 提交后需额外做一次 spot-check