# Excel 共享文件协同编辑方案（v2 — 列级声明 + 三方合并）

> v1 是「整表接管」；v2 按用户新方案改为「sheet + study 列级声明 + 三方合并」。
> 标注 **[待测]** 的结论依赖尚未取得可靠实测数据，实施前必须补齐。

---

## 0. v1 → v2 变更摘要

| 项 | v1 | v2 |
|---|---|---|
| 声明粒度 | sheet | **sheet + study 列集合** |
| 并发 | 同表只能一人 | **同表不同列可并发** |
| 合并 | 无（整表覆盖） | **三方合并 Base/Local/Remote** |
| 基线 | 无 | **必须有 Base 快照**（v1 缺失，v2 补上） |
| 发布方式 | 重命名换页 | **打开当前文件 + `Save()`**，两者均已排除/确定 |
| 按钮 | 1 个 | **4 个 Ribbon 按钮 + 状态机** |

---

## 1. 目标

单文件共享 Excel 的多人协同编辑：用户签出「某 sheet 的某组 study 列」→ 本地编辑 →
提交时三方合并回服务器文件，不覆盖他人改动。

## 2. 非目标

- 不提供增删行/列、排序、合并单元格
- 不支持多人编辑同一 study 列的并发合并（用锁串行化）
- 不做云端/SharePoint/OneDrive
- 不修复历史脏数据（幽灵行等）

---

## 3. 关键概念

### 3.1 MDDT 类型 —— ✅ 已决定：**不区分类型**

用户答复（2026-11-18）：`mddt_vlm` / `mddt_ct` / `mddt_sdtm` / `mddt_adam`
**在结构上没有区别**，第 1 行表头、前 5 列为锚点、AMG 开头的列是 study 列，
四类文件完全一致。

因此**文件名前缀不参与任何逻辑判断**：

| 原设计 | 现设计 |
|---|---|
| 按 `mddt_sdtm` / `mddt_adam` / `mddt_vlm` / `mddt_ct` 分支 | ❌ 删除所有分支 |
| 每类文件有各自的锚点列集合 | ✅ 统一用「前 5 列」 |
| 每类文件有不同的 sheet 结构 | ✅ 统一处理 |

**收益**：`MdlSchema` 不需要按类型分派，少一层复杂度，也少一类测试。
文件名前缀只在 UI 里用于**显示**（让用户知道自己在操作哪个文件），不参与规则。

> ⚠️ 前 5 列作锚点的约束见 §3.3。

### 3.2 sheet 与列的分类（已按用户答复定稿）

**表头位置：始终第 1 行。**

列分两类，**靠前缀判断，不靠位置**：

| 类别 | 判定规则 | 可写 |
|---|---|---|
| **product 列** | 表头**不以** `AMG` 开头（不区分大小写） | ❌ 只读 |
| **study 列** | 表头**以** `AMG` 开头（不区分大小写） | ✅ 可签出 |

命名形态（用户确认）：`AMG` + 任意后缀，例如 `AMG001_STUDYNAME`。

| 特征 | 实测/确认 |
|---|---|
| 前缀 | **固定 `AMG`** |
| 大小写 | **不固定** → 必须 `UCase$` 归一化后比较 |
| 后缀 | study 名字，**随机**，无固定格式 |
| 数量 | 「有很多 study 列」 |
| 非 study 列 | 靠「不以 AMG 开头」+ 位置识别（都在 AMG 列左边） |

**sheet 分类：**

- **study sheet**：第 1 行存在至少一个 `AMG*` 列
- **product sheet**：不存在任何 `AMG*` 列 → **不可签出**（整张表只读）

**已确认**：用户可以**多选若干个 AMG 列**组成一次签出。

→ 「study」= 本次签出中用户选定的 **AMG 列子集**（可任意组合）。
→ 锁粒度 = 「单个 AMG 列」，一次签出 = 获取 N 个列锁（见 §6.1）。

> 归一化后重名的处理见 §3.3.1。用户已确认业务上不会发生（`AMG` 后的 `xxx` 是 studyid，不重复）；
> 代码里保留 3 行断言，只为**拒绝静默错误**，前置条件成立时永不触发。

### 3.3 锚点 = product 列（上限 5）—— 业务语义保证，不做检测

用户答复（2026-11-18，第二轮澄清）：

> 「前五列是如果 Product 列有这么多列的情况下，**如果没有就用有的那些列**」

所以锚点的定义不是"前 5 列"，而是：

```vba
firstAMG  = 第一个以 AMG 开头的列的列号
nProduct  = firstAMG - 1              ' product 列数
nAnchor   = nProduct : If nAnchor > 5 Then nAnchor = 5
```

**只取 product 列**——因为 product 列是工具唯一不写的列，只有它才是稳定的行标识。
若把 AMG 列也当锚点，一旦有人改了该列，行就"找不到"了。

| product 列数 | 锚点列数 |
|---|---|
| 3 | 3 |
| 4 | 4 |
| 5 | 5 |
| 8 | **5**（截断） |
| 0 | **0**（无锚点，见下） |

### 3.3.1 两条业务语义前置条件（用户确认，不写检测）

用户明确答复：这两个问题在业务上**不可能发生**，因此不需要测试。

| # | 前置条件 | 用户的依据 |
|---|---|---|
| **P1** | 同一张表内，AMG 表头**不可能**大小写归一化后重名 | `AMGxxx` 中 `xxx` 是 studyid，不会重复 |
| **P2** | product 列组合**不可能**不唯一 | 这是业务语义，行天然唯一 |

**我们不写探测脚本**（原 T21 / T22 已作废）。

但两者在代码里的表现方式不同，必须区分：

#### P1 —— 保留 3 行运行时断言（不是测试）

`UCase$` 归一化是**必需**的（你确认过 `AMG` 前缀大小写不固定），
所以"两个不同写法归一化后相同"这个隐患是**由构造决定存在**的，与概率无关。

`MdlSchema` 用 `Dictionary` 以大写表头为键建索引。若发生重名：

| 不处理 | 后果 |
|---|---|
| `dict.Add` 被跳过 / 或覆盖 | **静默**只锁一列 → 两人同时改 → 锁失效 |

```vba
' MdlSchema 建列索引时
key = UCase$(Trim$(hdr))
If colIndex.Exists(key) Then
    Err.Raise vbObjectError + 2001, "MdlSchema", _
        "Duplicate AMG header after case fold: " & hdr & _
        " (sheet " & ws.Name & "). Refusing checkout."
End If
colIndex.Add key, colAddr
```

> 这不是"检测"（不扫描、不报告、不跑测试），而是**拒绝静默错误**。
> 前置条件成立时它永远不触发。

#### P2 —— 不写任何断言

唯一性若由业务语义保证，代码里检查它只是浪费与噪声。
**直接信任前置条件。**

#### 仍然保留的两项机制（属于机械问题，不是语义问题）

| 机制 | 为什么保留 |
|---|---|
| **跳过锚点全空的行** | 空行、分组标题行没有锚点。若参与合并，多行共享空键会互相塌陷。这是**机制**问题，不是业务唯一性问题 |
| **`nAnchor = 0` 时报错** | 若某表 AMG 列在第 1 列，就没有任何 product 列可作锚点。此时无法定位行，必须拒绝签出而不是猜 |

### 3.3.2 若前置条件被违反，症状是什么

写在这里，便于将来排查（不需要现在做任何事）：

| 被违反 | 症状 | 排查方向 |
|---|---|---|
| P1 | 两人声称锁了不同列，实际同时改同一列 | 查两张表头是否只差大小写 |
| P2 | 合并后某行的值出现在另一行 | 查 product 列组合是否有重复 |

### 3.4 改动的定义（判定白名单）—— ✅ 已定稿

用户答复：**加粗也算改动**（L2 默认开启）。完整白名单见下。

#### 单元格属性全清单（按 Excel 对象模型）

| 类别 | 属性 | 用户是否会改 |
|---|---|---|
| **内容** | `Formula` / `Value2` / `Text` | 是（主要） |
| | `NumberFormat` / `NumberFormatLocal` | 是 |
| | `PrefixCharacter`（前导单引号标记） | 间接 |
| | `RichText`（部分字符富文本） | 是 |
| **字体** | `Font.Name` / `Font.Size` | 是 |
| | `Font.Bold` / `Font.Italic` | **是（用户明确要求计入）** |
| | `Font.Underline` / `Font.Strikethrough` | 是 |
| | `Font.Color` / `Font.ColorIndex` | 是 |
| | `Font.ThemeColor` / `Font.TintAndShade` | 是 |
| | `Font.Superscript` / `Font.Subscript` | 是 |
| **填充** | `Interior.Color` / `ColorIndex` | 是 |
| | `Interior.Pattern` / `PatternColor` | 较少 |
| | `Interior.ThemeColor` / `TintAndShade` | 是 |
| **边框** | `Borders(6 边 + 对角线).LineStyle` | 较少 |
| | `.Weight` / `.Color` / `.ThemeColor` | 较少 |
| **对齐** | `HorizontalAlignment` / `VerticalAlignment` | 较少 |
| | `WrapText` / `ShrinkToFit` / `Orientation` / `IndentLevel` | 较少 |
| **保护** | `Locked` / `FormulaHidden` | 否（结构级） |
| **结构** | `MergeCells` / `MergeArea` | 否（本工具禁止合并） |
| **附加** | `Comment.Text` | 是 |
| | `Hyperlink.Address` / `SubAddress` / `TextToDisplay` | 是 |
| | `Validation.Type` / `Formula1` / `Formula2` | 较少 |
| | `FormatConditions`（条件格式） | 较少 |

#### 真实规模决定了实现路线（已由用户确认）

> **实际 study 列最多约 500 行，且一般只会选 3–4 列 → 约 1,500–2,000 格。**

因此之前担心的「几十个属性 × 15 万格 = 600 万次 COM 调用」**在真实场景不存在**。

实测参照（开发机，`VERIFIED.md` §12.3）：

| 规模 | 读 `.Formula` | 纯 VBA diff | 逐格读填充色（外推） |
|---|---|---|---|
| 2,000 格（真实规模） | < 0.01 s | < 0.01 s | **≈ 0.08 s** |
| 25,000 格（50 列上限） | 0.02 s | 0.01 s | ≈ 1.05 s |

→ **逐格读样式完全可行，不需要任何「一次拿全部属性」的奇技淫巧。**

#### ❌ 已否决的方案：`Range.Value(11)` XML

曾考虑用 `Value(11)`（`xlRangeValueXMLSpreadsheet`）一次拿到值+字体+填充+边框+对齐，
因为它确实能捕获这些（连**合并格、超链接、批注**都能，已验证）。

但实测发现 **它不是确定性的**：同一区域连续读两次相同，
但**保存重开后不同**、**颜色改回原值后也不同**——差异来自内部的 StyleID 计数器
（`63` vs `16`，见 `VERIFIED.md` §12.8）。

→ 拿它做指纹会**每次都报「全变了」**，产生海量假冲突。**已否决。**

#### ✅ 最终采用的分层白名单（已拍板）

| 层 | 包含 | 默认 | 理由 |
|---|---|---|---|
| **L1 必比** | 值 / 公式、数字格式 | ✅ 开 | 业务数据本体 |
| **L2 外观** | 填充色、字体色、**加粗**、斜体、字号、字体名 | ✅ **开** | 用户明确要求加粗计入；真实规模下成本可忽略 |
| **L3 细粒度** | 边框、对齐、下划线、删除线、上下标 | ⬜ 关 | 噪声大、业务价值低，可后续开启 |
| **L4 排除** | 行高、列宽、批注、数据验证、条件格式、合并 | ❌ 永不比 | 用户明确说行高列宽不重要 |

#### 颜色比较的已实测规则（T7）

| 规则 | 依据 |
|---|---|
| 用 `.Interior.Color` / `.Font.Color` 单值比较 | 已归一化为 RGB，主题色与同值显式色报同一个值 |
| **不**比较 `ThemeColor` / `TintAndShade` | 否则"同一颜色不同来源"会变成假冲突 |
| 保存/重开无颜色漂移 → 可安全启用颜色比较 | 实测 4 项属性全部一致 |
| **必须校验主题指纹** | 主题一换，已设的主题色填充会变值且换回来不恢复 |

#### ⚠ 必须警告：比得越多，假冲突越多

`Base` 取自**本地副本**，`Remote` 取自**服务器文件**。两者经过 Excel 不同次的
打开/保存后，部分属性会被「规范化」：

- 主题色解析成 RGB 的时机不同
- 未显式设置的属性回落到默认值（可能与源文件不同）
- `Value(11)` 的 XML 可能含易变的时间戳或内部 ID

→ 若「什么都比」而不设白名单，会产生**大量假冲突**，用户会直接失去信任。

→ **结论：必须用显式白名单 + 归一化，而不是「所有属性」。**

#### 未纳入白名单的属性 = 不追踪

若两个用户各自改了不同的未追踪属性，写入时后者覆盖前者，无人察觉。
这是**已知并接受的取舍**，需写进用户手册。

**不做事件追踪**（不挂 `Change` 事件）。原因：
- 事件会被 undo / 宏 / 粘贴 / 填充绕过或重复触发
- 崩溃后事件状态丢失
- 批量的 `Change` 事件会拖慢 Excel

**改为提交时整体 diff**：`Base` vs `Local` 逐格比较。
这样**天然支持任意编辑方式**（打字、粘贴、填充、拖拽、宏、undo）——
直接回答了「是否支持批量粘贴」。

#### ⚠️ 白名单已拍板，但「定义本身」仍有一个未决点

用户 2026-11-18 追问：「改动定义是不是也没确定？」

**答案：白名单（比哪些属性）已定稿，但「什么算冲突」还没定。**

这两件事必须分开：

| 层 | 状态 | 说明 |
|---|---|---|
| **比什么属性** | ✅ 已定稿 | L1/L2 默认开，L3 关，L4 永不比（上表） |
| **比出来的差异怎么判定** | ⬜ **未定** | 下面三种情形都需要明确规则 |

##### 未决点 D1：product 列被改了怎么办

现状是：工具只写回 AMG 列（study 列），product 列**只读**。

| 场景 | 若不处理 | 应该怎样 |
|---|---|---|
| 用户签出后，**手工改了 product 列** | 提交时不会被写回，但本地副本看着"改了" | 提交前**明确拒绝**："product 列不可修改，请撤销 D5 处的改动" |
| 服务器上别人的 product 列变了 | 工具不动它，正常 | 无需处理 |

→ 需要决定：**拒绝提交** vs **静默忽略** vs **警告后允许**。
→ 我倾向**拒绝**（静默忽略会让用户以为改生效了）。

##### 未决点 D2：同一列里「值没变但格式变了」

| 情形 | L2 是否算改动 |
|---|---|
| 只把某格从常规改成加粗 | ✅ 算（用户明确要求） |
| 只把填充色从无色改成白 | ✅ 算（实测两者 `.Color` 相同但 `Pattern` 不同 → 目前**不计入**） |
| 只改数字格式 `0.00` → `0.0000` | ✅ 算（L1） |

→ 白名单已覆盖这些，但**"算作改动"之后要不要提示用户**还没定。
→ 我倾向：**不必提示，直接写回**（否则每个格式变化都弹窗会很烦）。

##### 未决点 D3：三方合并的冲突判定粒度

这是最关键的一个，且**决定 `MdlMerge` 怎么写**：

| 粒度 | 规则 | 后果 |
|---|---|---|
| **A. 逐单元格** | 同一格 Base≠Local 且 Base≠Remote 且 Local≠Remote → 冲突 | 精确，但**同一行不同格**不会冲突 |
| **B. 逐行** | 同一行有人在任何格改了 → 整行冲突 | 保守，但会**误伤**（我改剂量、他改单位，本可共存） |
| **C. 逐列** | 同一列有人改了 → 整列冲突 | 太粗，锁已经是列级，这样等于没并发 |

→ **已锁定列级锁，所以列级冲突（C）永不会发生**（同一列同时只有一人持有）。

##### ✅ D1 已定：product 列被改 → **拒绝提交**

提交前检查所有 product 列 `Base ≠ Local`。发现则**整体拒绝**，指出位置：

```
┌────────────────────────────────────┐
│ 无法提交                            │
│ 以下 product 列不可修改：           │
│   AE!D5   (Base=CDISC, Local=PCDISC)│
│   AE!F12  (Base=Y,     Local=N)     │
│ 请撤销这些改动后重试。              │
│                        [ 确定 ]    │
└────────────────────────────────────┘
```

理由：静默忽略会让用户以为改动生效了，下次打开发现没了，不知道原因。

##### ✅ D2 已定：格式变化 → **不提示，直接写回**

L2 白名单内的变化（填充色/字体色/加粗等）视为正常改动，不弹窗。
否则每个格式变化都提示，用户会被噪声淹没。

##### ✅ D3 已定：冲突粒度 = **逐单元格（A）**

| 场景 | 判定 |
|---|---|
| 张三改第 100 行 `AMG01` 剂量，李四改第 100 行 `AMG02` 单位 | **两人都成功** |
| 同一格 `Base≠Local` **且** `Base≠Remote` **且** `Local≠Remote` | 冲突 |

**核心理由**：列锁已经隔离了并发。同一行不同列是两个**独立的列锁**，
报冲突属于额外制造的假冲突。逐行判定（B）会在改剂量 + 改单位这种
本可共存的情况下无故拦截，提交越频繁误伤越多。

`MdlMerge` 的判定函数（**最终形态**）：

```vba
' 返回 True 表示冲突
Function IsConflict(ByVal b, ByVal l, ByVal r) As Boolean
    ' 我没改 -> 永远不冲突，用服务器的值
    If Same(b, l) Then Exit Function
    ' 服务器没变 -> 我有权提交
    If Same(b, r) Then Exit Function
    ' 双方都改且改成一样 -> 一致，不算冲突
    If Same(l, r) Then Exit Function
    ' 双方都改且不同 -> 真冲突
    IsConflict = True
End Function
```

> 注意顺序：`Same(b,l)` 必须最先判断。若先判 `Same(b,r)`，
> 「我没改但服务器变了」会被误判成我有改动要提交。

---

## 4. 部署形态

- 单一 `.xlam` 加载项，含 `CustomUI14.xml`（Ribbon）
- 零运行时依赖（不用 Python/.NET/COM 注册、不需要管理员权限）
- `Scripting.Dictionary` 为 Windows 内置 COM（`scrrun.dll`），可用；
  若要求绝对零外部对象，改用 VBA 内建开放寻址哈希
- 宏已默认启用（`VBAWarnings = 1`，测试 1）→ **不需要签名证书**

### Ribbon 按钮（3 个）

| 按钮 | 行为 |
|---|---|
| `Select MDDT` | 弹窗选文件 → 按文件名前缀判定类型 |
| `Select Sheet & Study` | 打开选择窗体：选 sheet（只列出含 `AMG*` 的 sheet）→ **多选 AMG 列** → 确认编辑时长（默认 20 分钟） |
| `Save Changes` | 三方合并 + 写回（详见 §9） |

### 选择窗体上的按钮（**不是** Ribbon 按钮）

| 按钮 | 行为 |
|---|---|
| `Import` | 在「选择 sheet & study」窗体上。检查锁 → 建本地副本 → 解除本地 sheet 保护 → 只允许编辑所选 AMG 列 → 关窗体 |
| `Cancel` | 关窗体，不加锁 |

按钮启用状态由状态机驱动（`IRibbonUI.InvalidateControl`）。

---

## 5. 状态机

```
Idle
 └─ [Ribbon] Select MDDT ──────────────► MddtSelected
      └─ [Ribbon] Select Sheet & Study
             └─ 窗体打开（可反复改选 sheet / AMG 列，此刻仍不加锁）
                   └─ [窗体] Import ────► CheckedOut   (加锁 + 建本地副本)
                         └─ [Ribbon] Save Changes ──► Idle
                               └─ 崩溃 ──► 锁靠心跳 + 过期回收

任一阶段点 [窗体] Cancel / 直接关窗体 → 回到上一状态，不加锁
```

**约束**：一个 Excel 进程同时只允许一个签出。

---

## 6. 锁定设计

### 6.1 粒度

> **锁 = (sheet, 单个 AMG 列)**

一次签出 = **获取 N 个列锁**（N = 所选 study 列数）。

理由：用户可选 AMG 列的任意子集；按列加锁能正确处理**部分重叠**的签出
（例如 A 选 `AMG01,AMG02`，B 选 `AMG03` → 不冲突；B 选 `AMG02` → 冲突）。
若按「study 组」加锁，重叠判定会很复杂且容易错。

### 6.2 互斥原语：**原子 `MkDir`**

**不要用一个共享的 XML 做锁表**——多客户端同时改写会损坏。
改为：**每个列锁 = 一个目录**，靠 `MkDir` 的原子性做 CAS。

```
<MDDT 目录>/history/mddt_lock/
    <FILE>__<SHEET>__<COL>.lock/   ← 目录存在 = 已锁定
        owner.txt                  ← key=value 元数据
        heartbeat.txt              ← 每次心跳重写，用 mtime 判断存活
```

**命名规则（重要）**

| 规则 | 原因 |
|---|---|
| 必须含**文件名**，不能只放类型 | 同一目录下可能有多个同类文件（`mddt_sdtm_A.xlsx`、`mddt_sdtm_B.xlsx`），只放类型会互相锁错。文件名本身已带类型前缀，一次满足两个需求 |
| **统一转大写** | Excel 的 sheet 名与列名都不区分大小写；Samba 可能配置为大小写敏感 → 大小写不同的两个键会各自创建成功，产生**双重锁** |
| 非法字符替换为 `_` | `\ / : * ? " < > |` 在路径中非法；sheet 名可能含空格 |
| 总长度受控 | Windows `MAX_PATH` = 260。`<FILE>` 段做**截断 + 8 位哈希后缀**，例如 `MDDT_SDTM_AE_2024__a1b2c3d4__AMG01.lock` |

示例：

```
mddt_sdtm_AE_20240101.xlsx + sheet "AE" + 列 "amg01"
  → MDDT_SDTM_AE_20240101__AE__AMG01.lock
```

`owner.txt` 内仍保存**原始大小写的键**，供 UI 原样显示。

- 创建目录用 VBA `MkDir` → 映射 `CreateDirectory`，**已存在则报错**，即原子判定
- 失败 = 已被占用 → 读取 `owner.txt` 告知用户「谁、从何时、到何时」
- 释放 = `Kill` 内部文件 + `RmDir`
- 已实测 `mkdir` 可作为互斥原语（见 `VERIFIED.md`）

> **[待测]** 两台不同机器上 `MkDir` 是否仍然原子（测试 4，需同事配合）。

### 6.3 `owner.txt` 字段（回答「lock file 的字段」）

```
schema=1
kind=column-lock
mddt=mddt_sdtm_xxx.xlsx
sheet=AE
column=AMG01
user=zhangsan
domain=CORP
machine=PC-1234
rdp_session=RDP-Sxs...#0
start_utc=2026-01-20T08:15:03Z
lease_seconds=1200
expires_utc=2026-01-20T08:35:03Z
heartbeat_utc=2026-01-20T08:24:41Z
base_fingerprint=9F3A...  (study 块内容指纹，见 §8.2)
tool_version=1.0.0
```

历史/审计：**不要**写共享 CSV（并发追加会损坏）。
改为每次事件一个文件：`history/mddt_lock/log/<UTC>_<user>_<sheet>_<COL>_<event>.txt`，
`event ∈ {acquire, renew, release, expire, steal}`。写入即新建，无竞争。

### 6.4 心跳与过期（回答「超时自动释放」）

#### 谁维护心跳（用户提问）

**由持锁者自己的 Excel 插件维护**，不需要任何服务端进程、不需要定时任务。

| 机制 | 说明 |
|---|---|
| `Application.OnTime` 每 60 s | 主心跳。只要 Excel 进程存活就触发 |
| 用户事件兜底 | `Workbook_SheetChange` / `SheetSelectionChange` 顺带刷新（成本≈0） |
| 显式操作时刷新 | 提交、续期、保存等动作前先刷一次 |

> `Application.OnTime` 的前提是：Excel 进程存活、插件仍在内存、
> **没有阻塞型模态对话框**。所以它只是**存活证明**，不是强保证；
> 真正的兜底是过期回收。
>
> 工作组旁路工具（§13）是**只读**的，**不维护心跳**。

#### 过期判定

```
过期 = (now > expires_utc)
    或 (now - heartbeat_mtime > 3 分钟 + 宽限)
```

- 主判据用 **`heartbeat.txt` 的服务器 mtime**（不受客户端时钟影响）
- `expires_utc` 仅用于**展示**给用户
- 过期后由**下一个使用者**顺手清理并写 `expire` 日志（无需定时任务）

#### 被抢占怎么办

若持锁者心跳停了超过 3 分钟，别人可正常抢占。原持锁者回来提交时：

1. 校验锁目录 `owner.txt` 是否仍是自己 → 不是 → **已失去锁**
2. 提示「你的签出已过期并被他人接管」
3. 但仍允许走**三方合并**提交（`Base` 快照还在本地）→ 冲突按 §8 正常裁决

→ **数据不会丢，只是要走冲突流程。**

#### 续期

- **到期前 5 分钟**提示「是否续期」，避免 20 分钟把还在干活的人踢掉
- 续期 = 改 `owner.txt` 的 `expires_utc` 并写 `renew` 日志

### 6.5 崩溃恢复（回答「任意退出都要释放」）

**诚实结论：Excel 崩溃时无法执行任何 VBA，因此"必然释放"做不到。**

分层兜底：

| 场景 | 兜底机制 |
|---|---|
| 正常关闭 / 点「放弃」 | `Workbook_BeforeClose` + 显式释放 |
| 用户忘记关闭就下班 | 心跳停止 → 他人按 §6.4 判定过期并清理 |
| Excel 崩溃 | 同上（心跳停止） |
| 进程被杀 | 同上 |
| 插件启动 | 启动时清理**本机本用户**的遗留列锁 |

**remote 文件句柄**：由操作系统在进程死亡时自动释放，无需担心。

### 6.6 时钟偏移

`expires_utc` 依赖客户端时钟。缓解：
- 取 `owner.txt` 写完后回读的 **服务器 mtime** 作为参考基准，而非纯本地时钟
- 或用「心跳 mtime」作为主要判据（mtime 由服务器产生，不受客户端时钟影响）
- **推荐以心跳 mtime 为主判据**，`expires_utc` 仅作展示

---

## 7. 本地工作区

```
%LOCALAPPDATA%\MDDTLock\<fileId>\
    working.xlsx      ← 本地副本（只含签出的 sheet，或整本副本）
    base.tsv          ← 签出瞬间 study 列的基线快照（三方合并必需）
    meta.txt          ← 文件路径、sheet、study 列、锁信息、指纹
    heartbeat.txt     ← 本地心跳（配合远端锁）
```

`fileId` = `文件名 + 规范化 UNC 路径 + 大小 + mtime` 的哈希。

**硬性要求**：**绝不能**放在 `Documents` / `Desktop`。
该环境 OneDrive 已重定向（实测），放那里会把业务数据同步上云。

---

## 8. 三方合并

### 8.1 为什么必须有 Base（对用户方案的关键修正）

用户方案写的是「Remote / Local copy / Changes」三方。
但 **`Changes` 不是独立对象**，它由 `Local` 与某个基线比较得出。
真正需要的三个对象是：

| 对象 | 含义 |
|---|---|
| **Base** | **签出瞬间**服务器上该 study 列的内容（存本地） |
| **Local** | 用户当前本地内容 |
| **Remote** | 提交瞬间服务器上该 study 列的内容 |

若没有 `Base`，`Local ≠ Remote` 无法区分「我改的」「他改的」「两人改成一样」，
只能对每一处差异都问用户 —— 不可用。

**所以：签出时必须把 Base 落到本地。** 这是 v1 的缺失，v2 必须补。

### 8.2 逐格三方合并规则

对每个 `(row, col)`：

```
mine   = (Local  != Base)
theirs = (Remote != Base)

if !mine && !theirs:  result = Base        # 都没动
elif mine && !theirs: result = Local       # 只我动
elif !mine && theirs: result = Remote      # 只他动
elif Local == Remote: result = Local       # 两人改成一样，不算冲突
else:                 result = CONFLICT    # 真冲突
```

这是标准三方合并。**冲突只在两种情况出现**：

1. 用户锁过期后，同一列被别人签出并提交
2. 有人绕过工具直接改了文件

由于锁是列级串行化的，冲突属于**异常路径**，所以冲突 UI 可以做得简单。

### 8.3 冲突 UI

- 汇总：`我的改动 N 处 / 对方改动 M 处 / 冲突 K 处`
- 冲突表：`sheet | 行 | 列 | Base 值 | 我的值 | 对方的值 | [用我的] [用对方的]`
- 提供「全部用我的 / 全部用对方的」批量按钮
- **默认不写入任何冲突格**，未裁决则禁止提交

### 8.4 快速路径（回答「MD5 是否更快」）

**不要用 MD5。**
- VBA 无内建 MD5，要自己实现（约 200 行），且比 FNV 慢
- 已实测并交叉验证的 **FNV-1a 双段 64 位**（`VERIFIED.md` §7）直接复用

**哈希的正确用处不是"比较"，而是"存指纹跳过工作"**：

```
签出时：fingerprint = FNV64(study 块内容)  → 写进 owner.txt / meta.txt
提交时：若 FNV64(当前 Remote study 块) == fingerprint
        → Remote 完全没被改过 → 不可能有冲突 → 直接全量写入
```

这样 90% 的提交不进入逐格比较。

> **注意**：读单元格的开销远大于哈希本身（COM 调用是瓶颈）。
> 所以「用哈希代替比较」并不能省掉读取。真正省掉的是**逐格分支判断**和**冲突扫描**。
> 结论：哈希用于**判定是否需要细比**，不是用于替代读取。

### 8.5 `Scripting.Dictionary` 是否够快（回答「dict 方法」）

规模估算：study 列通常 ≤ 30 列 × 行数（常规几十到几千行）。
即使 30 × 5000 = 15 万格，`Dictionary` 也足够（毫秒级）。

**瓶颈不在容器，在 Excel COM 调用次数。** 因此：
- 值：一次 `Range.Formula` 拿 2D 数组（1 次 COM 调用）
- 填充色 / 字体色：`Range.Interior.Color` 对多格区域在颜色不一致时返回 `Null`，
  **必须逐格读** → 这是主要开销
- 优化：默认只比内容；仅在用户勾选「包含颜色」时才逐格比色。
  或对整块取 `Range.Value(xlRangeValueXMLSpreadsheet)`（一次调用返回含样式的 XML）
  后做字符串哈希 —— **[待测]** 其在本文件上的耗时

---

## 9. 提交流程（核心事务）

```
[Save Changes]
 1. 延迟预检：5 次极小 IO，测平均延迟
       > 500 ms → 警告「共享盘严重拥塞，建议稍后再提交」，允许强行继续
 2. 校验本地锁仍有效（心跳未过期）；无效 → 提示并禁止提交
 3. 对目标 (sheet, 列) 逐个 MkDir 获取**提交锁**（他人签出中的列则拒绝）
    —— 提交前**不做**内容判定用的 `wb.Saved` 检查（T18：脏信号不可靠）
 4. 在**宿主实例**中打开 Remote（read-write），全程 `EnableEvents=False`
    ← 唯一占用服务器文件的时刻（**不新建 Excel 实例**，见 §12.2）
    4b. **立即** `Remote.Windows(1).Visible = False` ← 焦点毫秒级归还（§12.3）
 5. 结构守卫（全部通过才继续，任一失败即中止，不写入）
       a. sheet 存在
       b. 第 1 行列名与签出时一致
       c. study 列仍位于相同列号
       d. 行数一致
       e. 每行锚点 = 声明 key 列集合的哈希，与 Base 逐行一致
 6. 三方合并计算（§8），得到 out[][] 与 conflicts[]
 7. conflicts 非空 → 关闭 Remote（不保存）→ 弹冲突 UI（在内存里裁决后重跑 3~6）
 8. Clear 目标区域并集 → Copy 写入（见 §9.1）
 9. 若 sheet 受保护：`Unprotect pwd` → 写 → 按原标志位 `Protect pwd`
10. `wb.Save()`  ← 一次完整写入，绝不部分保存
11. `wb.Close`
12. 释放提交锁 + 签出锁，写日志，清理本地工作区
```

### 9.1 写入必须用 `Range.Copy Destination:=`

已实测：
- `Value2` 数组写入 → **超链接 0/2、填充色丢失**
- `Range.Copy` → 21/21 项保真（超链接含内部 SubAddress、填充色、字体色、
  数字格式、边框、合并格、批注、数据验证、条件格式）
- **`Range.Copy` 不携带行高/列宽** → 需显式赋值

### 9.2 必须「先 Clear 并集再写」

已实测：只覆盖较窄区域会残留旧行，表现为「删除失败」。

### 9.3 结构守卫的锚点（对用户方案的修正）

用户方案是"对比 product 列前五列的值"。

**改进**：anchor = **声明的 key 列集合**的逐行哈希，而不是"前五列"。

理由：
- "前五列"在没有 key 的 sheet 上可能全是空列或说明列 → 无判别力
- key 列集合可按文件类型配置，比固定 5 列更可靠
- 若 key 集合唯一 → 还能支持**按键匹配行**，容忍行移动

**三种文件类型的锚点策略不同**（用户已确认）：

| 类型 | 锚点 |
|---|---|
| `mddt_sdtm` / `mddt_adam` | 大部分 sheet：key 列组合唯一 → 用 key 匹配行 |
| 部分 sheet | 无唯一组合，但模板只允许加列 → 用**行号** + key 哈希校验 |
| `mddt_vlm` / `mddt_ct` | **[待确认]** 结构差异，实施前逐类确定 |

---

## 10. 失败与恢复

> 用户判断正确：**风险只存在于保存那一步**。
> 保存前的所有操作都在内存或本地，失败不影响远程文件。

| 失败点 | 后果 | 处理 |
|---|---|---|
| 提交前任何一步 | 远程文件未被触碰 | 直接重试或放弃，本地数据完好 |
| 打开 Remote 时被他人独占 | 未修改任何内容 | 提示「文件被占用，稍后重试」 |
| 结构守卫失败 | 未写入 | 中止并说明原因（如"检测到新增行"） |
| `Save()` 失败 | ⚠️ 风险点 | 见下 |
| Excel 在 Save 期间崩溃 | ⚠️ 风险点 | 见下 |

### 10.1 保存阶段的保护

Excel 保存 xlsx 时**在同一目录写临时文件再改名**（这是 Excel 自身行为），
因此理论上失败不会破坏原文件。**[待测]** 需确认该行为在本 Samba 共享上成立。

补充措施：
- **不做**「提交前把 11 MB 原文件复制成备份」——
  新建文件在本共享上有 106–245 s 惩罚（实测），代价过高
- 改为：保留**本地**工作副本与 Base 直到用户确认提交成功；
  任何失败都可重新提交
- 提交失败后**不自动重试**（避免在异常网络下反复写入），只提示用户重试

### 10.2 幂等性

提交逻辑设计为**幂等**：以 `Base` 为参照重新计算一次即可安全重放。
因此「失败 → 重新点 Save Changes」是安全的。

---

## 11. 模块架构（解耦）

```
MdlInfra       Win32 Declare / 路径与 UNC 解析 / 日志 / 错误包装 / 延迟探测
MdlStorage     %LOCALAPPDATA% 工作区、base 快照读写、fileId、原子改名
MdlLock        MkDir 互斥、owner.txt、心跳、过期清理、提交锁
MdlSchema      文件名→类型、sheet 分类、AMG* 列解析、key 列配置、列名指纹
MdlExcel       **唯一**直接接触 Excel 对象模型的模块：
               Open/Close/Save、Range 读写、Copy、保护开关、行高列宽
MdlDiff        Formula/Color 采集、FNV-1a 64 哈希、逐格比较
MdlMerge       三方合并、冲突收集、out[][] 生成
MdlSession     状态机、单签出约束、按钮启用
MdlUI          Ribbon 回调、选文件/选 sheet+study 窗体、冲突裁决窗体
```

规则：
- 除 `MdlExcel` 外，任何模块都**不得**出现 `Excel.Application`、`Range`、`Worksheet`
- `MdlMerge` / `MdlDiff` 只接受内存数组 → 可脱离 Excel 单元测试
- `MdlUI` 不直接调用 `MdlLock` / `MdlMerge`，只经 `MdlSession`

---

## 12. 性能目标（已实测）

> **数据来源**：测试 7（公司内网，VBA，N=20，0 失败）。
> 之前的外网 WAN 数据已全部作废，不再引用。

### 12.1 提交窗口（真实值）

| 阶段 | avg | p95 | max |
|---|---|---|---|
| 打开服务器文件 | 3.27 s | 3.88 s | 3.97 s |
| 写入声明列 | < 0.1 s | — | — |
| `Save()` 回共享盘 | 6.77 s | 7.41 s | 8.48 s |
| `Close()` | < 0.2 s | — | — |
| **合计** | **9.41 s** | **10.56 s** | **10.62 s** |

→ 稳定性极高（min 8.65 / max 10.62）。比早期推算的 20–25 s 更好。

### 12.2 实例策略（已决策）

| 方案 | 窗口 avg | 结论 |
|---|---|---|
| **宿主 Excel 实例**（`Application.Workbooks.Open`） | **9.41 s** | ✅ **采用** |
| `CreateObject` 新实例常驻 | 12.64 s | ❌ +34% |
| 每次提交都新建实例 | 17.09 s | ❌ +82% |

额外发现：新实例连 Open/Save 本身也更慢（4.76 vs 3.27；7.73 vs 6.77），
即使文件已被预热。→ 不是缓存问题，是**冷进程固有开销**。

→ **无性能理由新建实例。**

### 12.3 宿主实例的四个副作用：已全部实测，其中三个可完全消除

T15 / T16 / 宿主机集成实验（`VERIFIED.md` §12.1、§12.2、§13）的结论：

| 原担心的副作用 | 实测结论 | 对策 |
|---|---|---|
| 触发插件自己的 `ThisWorkbook.Workbook_Open` / `BeforeClose` | **不会触发**。candidate 的 open/close 只产生 `Application` 级事件 | 释放锁的逻辑挂在 `ThisWorkbook`，不要挂 `Application` 级 |
| 触发 `Application` 级事件导致递归 | **完全归零**。`EnableEvents=False` 下事件数为 0 | `EnableEvents=False` + `mInCommit` 双保险 |
| 抢焦点、切工作簿、界面乱跳 | **可消除**。`Open` 后立即隐藏窗口，活动工作簿/表/选区立刻完整归还 | `ScreenUpdating=False` + `Visible=False` |
| 提交期间用户继续操作 | 用**模态进度窗体**占住 UI | 窗体上显示阶段与提示 |
| 用户已把目标文件开着 | **约定范畴**，工具不为此设计额外流程 | 见 §14 Q10 |

#### 焦点归还的实测证据

| 时刻 | 活动工作簿 | 活动表 | 选区 |
|---|---|---|---|
| 用户态 | Book3 | Sheet1 | **D5** |
| 刚 `Open` 完 | **candidate**（被抢） | AE | A1 |
| **`Visible=False` 后** | **Book3** | **Sheet1** | **D5**（立刻回来） |
| close 后 | Book3 | Sheet1 | D5 |

窗口暴露时间 = 几毫秒，且 `ScreenUpdating=False` 下无重绘，此时用户还被模态窗体挡着。
**用户的选区也没有丢失。**

隐藏后所有操作照常可用：`Range.Copy Destination:=`（隐藏→隐藏）正常、
`wb.Save` 正常、`Range.Clear` 正常。
→ **整个提交流程可以在一个不可见的 workbook 上完成。**

#### 不使用 `Application.Interactive = False`

- 实测**不能**让新打开的窗口自动隐藏（仍然 `Visible=True`）
- 实测**没有**任何额外收益（焦点归还序列逐项相同）
- 若异常退出时卡在 `False`，用户整个 Excel 会**失去响应** ——
  比"焦点闪一下"严重得多

#### 规范写法

```vba
Private mInCommit As Boolean

Sub Commit()
    Dim oldEE As Boolean, oldSU As Boolean
    oldEE = Application.EnableEvents
    oldSU = Application.ScreenUpdating

    mInCommit = True
    On Error GoTo Cleanup

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(remotePath, 0, False)
    wb.Windows(1).Visible = False       ' 焦点立刻归还

    ' ... 结构守卫 / 三方合并 / 逐列 Copy / wb.Save ...

    wb.Close False
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = oldSU
    Application.EnableEvents = oldEE
    mInCommit = False
End Sub
```

**两条铁律**：

1. 所有 `Application` 级事件处理第一行必须是 `If mInCommit Then Exit Sub`
2. `EnableEvents` / `ScreenUpdating` / `Calculation` 的恢复必须在 `Cleanup` 里，
   且 `Cleanup` 内部用 `On Error Resume Next` —— 否则一次异常会把用户的全局设置卡死
   （公式不再自动重算，是最容易被投诉的 bug）

#### 剩余风险

| 风险 | 等级 | 说明 |
|---|---|---|
| 提交中途异常后全局 `EnableEvents` 卡在 False | 中 | 靠 `Cleanup` + `Resume Next`；建议启动时自检修复 |
| candidate 是宏文件（`.xlsm`/`.xlam`） | 低 | 本工具只开 `.xlsx`；可加 `AutomationSecurity = ForceDisable`（T17 未测） |

### 12.4a Save 的原子性：不需要提交前备份（已实测，`VERIFIED.md` §16）

监控保存期间目录变更，抓到的真实序列：

```
 7.1s  CREATE   55790410           ← Excel 把新内容写进临时文件
 9.6s  MODIFY   55790410           （原文件此期间完全未被触碰）
 9.649 CREATE   ACCB5A11.tmp       ← 原始内容回滚副本
 9.649 DELETE   big_probe.xlsx     ┐ 同一毫秒
 9.649 CREATE   big_probe.xlsx     ┘ rename 语义，不是就地覆写
```

**杀进程实测（13.81 MB 探针，Save 耗时 4.2–5.8 s）：**

| 杀点 | 是否落在 Save 窗口内 | 结果 |
|---|---|---|
| @500 ms | 是 | zip 完好、Excel 能开、内容是**旧内容** |
| @1500 ms | 是 | zip 完好、Excel 能开、内容是**完整新内容** |
| 保存后杀（对照） | — | 完好 |

→ **`Save()` 被打断不会损坏服务器文件。** 最坏情况是"改动没生效"。

#### ★ 已在真实 SMB 共享上复验（`VERIFIED.md` §16.7）

本地盘的结论必须复验，因为 Excel 对网络位置可能走不同策略。真实 `book.xlsx`（10.94 MB）结果：

| 阶段 | zip 完好 | Excel 能开 | 杀点落在窗口内 |
|---|---|---|---|
| P1 正常保存 | ✅ | ✅ | — |
| P3 杀 @500 ms | ✅ | ✅ | **是** |
| P4 杀 @1500 ms | ✅ | ✅ | **是** |
| P5 杀 @3000 ms | ✅ | ✅ | **是** |
| P6 杀 @5000 ms | ✅ | ✅ | **是** |
| P7 保存后杀 | ✅ | ✅ | 否 |

四轮真正命中窗口的杀进程，**没有一轮留下损坏文件**。文件大小波动仅 25 字节（zip 时间戳）。

→ **本地盘结论可直接推广到生产路径，"不做提交前备份"成立。**

#### ⚠️ 新增：清理孤儿 `.tmp`（T20）

共享上崩溃（保存完成但清理未执行时被杀）会残留 `52C6A2C6.tmp` 这类文件：

| 问题 | 结论 |
|---|---|
| 损坏文件 | 不会 |
| 阻塞下次保存 | 不会（随机名字，不撞名） |
| 累积 | **会**，每次崩溃留 1–2 个 |
| 被误认为数据文件 | 不会（扩展名不是 `.xlsx`） |
| 被 SAS 通配符扫到 | **风险点** |

```vba
' 只删"老"的 8 位十六进制 .tmp，绝不删当天的（可能是别人正在保存）
Function IsOrphanTmp(ByVal fname As String, ByVal mtime As Date) As Boolean
    If Len(fname) <> 12 Then Exit Function                 ' 8 + ".tmp"
    If StrComp(Right$(fname, 4), ".tmp", vbTextCompare) <> 0 Then Exit Function
    Dim i As Long, c As String
    For i = 1 To 8
        c = Mid$(fname, i, 1)
        If InStr("0123456789ABCDEFabcdef", c) = 0 Then Exit Function
    Next i
    IsOrphanTmp = (DateDiff("h", mtime, Now) > 24)
End Function
```

#### 三条落地规则

| # | 规则 | 依据 |
|---|---|---|
| 1 | **不做提交前备份** | rename 保证原文件在写入期间不被触碰。省下 106–245 s |
| 2 | **失败恢复 = 重新打开核对，不是修复，也不是假设成功** | `0x800706BE` 后状态未知 |
| 3 | `Save()` 抛错时**绝不释放锁** | 否则他人可能读到未提交状态 |

```vba
' 正确：Save 失败后的处理
On Error Resume Next
wbServer.Save
If Err.Number <> 0 Then
    Dim e As Long: e = Err.Number
    Err.Clear
    On Error GoTo 0
    ' 状态未知 -> 重新打开核对，锁保持
    Select Case VerifyServerState(wbServerPath, expectedCols)
        Case VS_OK:      GoTo Done          ' 其实成功了
        Case VS_STALE:   GoTo Retry         ' 没写进去，安全重试
        Case VS_UNKNOWN: Abort "保存状态不明，请人工核对"; GoTo Cleanup
    End Select
End If
```

> 注意：Windows 侧轮询分辨率约 1 ms，rename 的瞬时"文件不存在"窗口
> 可能被漏掉。**Linux/SAS 侧是否读到 `ENOENT` 尚未验证（T5）**，
> 若成立则需在 SAS 消费方加读取重试，而不是改 Excel 侧。

### 12.5 角色识别与插件级状态（已实测，`VERIFIED.md` §14）

> 用户提出的方向：严格识别 workbook 角色 + 增加插件级状态，
> 让 `Application` 事件处理器按角色分流。**方向正确，但实现要点与直觉相反。**

#### ❌ 三个"直觉写法"实测是错的

| 直觉写法 | 实测结果 |
|---|---|
| 遍历 `Application.Workbooks` 找 `Is ThisWorkbook` 来识别插件自身 | **找不到**。插件工作簿不在 `Workbooks` 集合里（`count = 0`） |
| 用 `If mWb Is Nothing Then` 判断本地副本是否还开着 | **错**。关闭后 `Is Nothing` **仍返回 False**，且不抛错 |
| 关闭后访问 `mWb.Name` | 抛 `-2147221080`，但 `Is` 比较不抛错 → **静默错判** |

**最危险的是第二条**：误以为本地副本还在，`Is` 比较永远 False，
于是所有事件都被当成"别人的" → 逻辑静默走错分支，**不报错**。

#### ✅ 正确的四层机制

| 层 | 机制 | 用途 | 依据 |
|---|---|---|---|
| **L0 闸门** | `EnableEvents=False` + `mInCommit` | 提交期间所有事件归零 | §13.3 实测归零 |
| **L1 粗筛** | `Application` 级 `WithEvents`，按角色分流 | 统一入口、错误包装、状态推进 | §13.2 |
| **L2 精确** | 给本地副本单独挂一个 `WithEvents` 类实例 | 只收本地副本事件，处理器里无需再判断是谁 | §14.2：**严格按角色隔离，互不串扰** |
| **L3 身份** | `FullName` 路径字符串 + `CustomDocumentProperty` | 跨会话识别 | §14.3 / §14.4 |

`Application` 级和 workbook 级的 `WithEvents` **是两个独立订阅**，
一个 `Cancel` 不会阻止另一个被调用 → **关闭拦截逻辑只在一处实现**，否则重复提示。

#### 角色枚举与识别函数

```vba
Public Enum EWB
    WBR_Unknown = 0        ' 别人的
    WBR_LocalCopy = 2      ' 我们签出的本地副本
    WBR_RemoteTarget = 3   ' 提交时从服务器打开的
    WBR_UserOther = 4      ' 用户自己的其他工作簿（绝不干预）
End Enum

' 只用路径字符串，不持有对象引用
Function RoleOf(ByVal wb As Workbook) As EWB
    Dim p As String
    On Error Resume Next
    p = wb.FullName                       ' 失效引用会在此抛错
    If Err.Number <> 0 Then
        Err.Clear: RoleOf = WBR_Unknown: Exit Function
    End If
    On Error GoTo 0

    If StrComp(p, mSess.LocalPath, vbTextCompare) = 0 Then
        RoleOf = WBR_LocalCopy
    ElseIf StrComp(p, mSess.RemotePath, vbTextCompare) = 0 Then
        RoleOf = WBR_RemoteTarget
    Else
        RoleOf = WBR_UserOther
    End If
End Function

' 存活判定：必须遍历 Workbooks 比 FullName，不能用 Is Nothing
Function IsAlive(ByVal wb As Workbook) As Boolean
    Dim i As Long, target As String
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    target = wb.FullName
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    For i = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(i).FullName, target, vbTextCompare) = 0 Then
            IsAlive = True
            Exit Function
        End If
    Next i
End Function
```

#### ⚠️ 「脏」不是可靠信号（T18 实测，`VERIFIED.md` §15）

给本地副本打角色标记 `CustomDocumentProperty` 会**把 workbook 弄脏**：

| 操作 | 是否弄脏 |
|---|---|
| `CustomDocumentProperties.Add` | **是** |
| 改已有属性的 `.Value` | **是** |
| `.Delete` | **是** |
| 读属性 / 读 `Count` | 否（完全惰性） |

**后果：`wb.Saved` 不能再用来判断"用户改过数据"。**

因为打标记本身就弄脏了它。四条落地规则：

| # | 规则 | 原因 |
|---|---|---|
| 1 | **绝不用 `wb.Saved` 决定"要不要提交"** | 打标记会弄脏，信号被污染 |
| 2 | 判定"用户改了什么"**只能靠 Base 快照 diff** | 这本来就是三方合并的核心 |
| 3 | 打标记后**立即 `Save`** | 否则属性丢失，且用户看到"未保存" |
| 4 | 生产代码**禁用** `BuiltinDocumentProperties` 造脏 | 它改的是真实元数据，会污染文件 |

**污点标记的正确用法**：只作**跨会话兜底识别**，不作内容判定。

```vba
Function RoleMark(ByVal wb As Workbook) As String
    Dim s As String
    On Error Resume Next
    s = CStr(wb.CustomDocumentProperties("MDDT_ROLE").Value)
    If Err.Number <> 0 Then s = "": Err.Clear
    On Error GoTo 0
    RoleMark = s
End Function

Sub SetRoleMark(ByVal wb As Workbook, ByVal v As String)
    On Error Resume Next
    wb.CustomDocumentProperties("MDDT_ROLE").Value = v    ' 已存在则直接改
    If Err.Number <> 0 Then
        Err.Clear
        wb.CustomDocumentProperties.Add "MDDT_ROLE", False, 4, v
    End If
    On Error GoTo 0
End Sub
```

> 重复 `Add` 同名属性会失败（`-2147467259`），所以必须先改 `.Value` 再退回 `Add`。
> 属性经 `SaveAs` / `SaveCopyAs` / 重开都能保留（T18e），只是**不能反推内容是否改过**。

**首选识别手段仍然是路径字符串**（§12.5 的 `RoleOf`），零副作用；
属性标记只用于崩溃后重新打开本地副本时确认"这是我们自己的副本"。

#### 插件级状态机

```vba
Public Enum ESess
    SS_Idle = 0
    SS_TargetSelected = 1
    SS_CheckingOut = 2
    SS_CheckedOut = 3
    SS_Committing = 4
    SS_Conflict = 5
    SS_Aborting = 6
End Enum

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error GoTo H                              ' 规则 1
    If mInCommit Then Exit Sub                   ' 规则 2（L0）
    If mSess.State = SS_Committing Then          ' 状态
        Cancel = True
        Exit Sub
    End If
    Select Case RoleOf(Wb)                       ' 规则 3（L1）
        Case WBR_LocalCopy
            HandleLocalClose Wb, Cancel
        Case Else
            ' 用户的其它工作簿：绝不干预
    End Select
    Exit Sub
H:
    LogError "App_WorkbookBeforeClose", Err.Number, Err.Description
End Sub
```

#### 七条铁律（全部来自实测）

| # | 规则 | 依据 |
|---|---|---|
| 1 | 每个处理器第一行 `On Error GoTo` | 处理器抛错会中断用户操作 |
| 2 | 第二行 `If mInCommit Then Exit Sub` | `EnableEvents` 归零 + 双保险 |
| 3 | 只对**确认属于我们**的角色 `Cancel = True` | `Cancel` 很有效，用错会把用户文件锁死 |
| 4 | 不遍历 `Workbooks` 找插件自身 | 它不在集合里 |
| 5 | 不靠 `Is Nothing` 判断存活 | 关闭后仍返回 False |
| 6 | 跨会话只用路径字符串识别 | 对象引用会失效且**静默错判** |
| 7 | 关闭拦截只在一处实现 | Application 级与 workbook 级是独立订阅 |

### 12.4 修订后的目标（真实规模）

> **用户确认：study 列最多约 500 行，一般只选 3–4 列。**
> 即典型 1,500–2,000 格，上限 500 × 50 = 25,000 格。

| 指标 | 目标 | 实测 / 依据 |
|---|---|---|
| 典型规模 | 500 行 × 3–4 列 ≈ 2,000 格 | 用户确认 |
| 上限规模 | 500 行 × 50 列 = 25,000 格 | 保守上限 |
| 读值（典型） | < 0.01 s | 0.143 s @ 150k → 线性外推 |
| 读样式（典型，逐格） | **≈ 0.08 s** | 0.21 s @ 5k 实测 |
| 读样式（上限 25k） | ≈ 1.05 s | 同上外推 |
| 写回（典型，逐列 Copy） | **< 0.1 s** | 2.673 s @ 150k |
| 打开服务器文件 | ≤ 4 s | 实测 3.27 s |
| `Save()` | ≤ 8 s | 实测 6.77 s |
| **提交窗口** | **≤ 15 s**（p95） | 实测 9.41 avg / 10.56 p95 |
| **超时** | **120 s** | 实测 max 10.62 s，留 10 倍余量 |

→ **性能不再是任何风险点。** 打开+保存占窗口的 99% 以上，
   diff 与写回都在 100 ms 量级。
→ 因此 **L1+L2 全部属性默认开启**，不需要“为了性能只比值”的退让。

> 仍保留提交前**延迟预检**（§9 第 1 步）：若用户居家办公 / 走 VPN，
> 预检会先拦住并提示，而不是让用户对着进度条干等。

---

## 13. 工作组的旁路工具

一个**只读**的小工具（脚本或按钮），回答一个问题：

> **「这个文件现在有被签出吗？」**
> 列出 `history/mddt_lock/` 下所有未过期的列锁：谁、哪个 sheet、哪些列、到何时。

用途：工作组要做**增删行列**时，先确认没有签出中的锁。
有锁时增删行列会让签出者的锚点校验失败 → 对方的提交会被中止。

> **[待确认]** 用户原话「被签出时可以直接使用原始文件加减行列等行为」
> 含义存疑。按上下文理解为：**未签出时才允许改结构**；
> 若确实需要"被签出时也能改"，则必须让签出者重新签出（Base 失效），
> 见 §14 Q7。

---

## 14. 待确认（阻塞实施）

| # | 问题 | 影响 |
|---|---|---|
| ~~Q1~~ | ~~"study" 粒度~~ → **已确认：可多选任意 AMG 列子集** | ~~锁粒度~~ 已定：按单个列加锁 |
| ~~Q9~~ | ~~改动白名单的最终范围~~ | ✅ 已定稿（§3.4 L1/L2 开，L3 关，L4 不比） |
| Q2 | AMG 列的命名规则？是否有 `AMG01..AMGnn` 连续编号？大小写实际形态？ | 列解析 |
| Q3 | 三种文件类型各自的 key 列集合（哪个 sheet 用哪些列做锚点） | 结构守卫 |
| Q4 | `mddt_vlm` / `mddt_ct` 的结构与 sdtm/adam 差异 | 锚点策略 |
| Q5 | sheet 保护：项目组能否改为「只锁 product 列、放开 AMG 列」？（见 §15.1） | 是否需要密码 |
| Q6 | 若不能改模板，42 张 sheet 的密码如何分发给用户？ | 方案 A/B/C 选型 |
| Q7 | 工作组在有人签出时能否改结构？ | 恢复策略 |
| Q8 | 是否允许筛选 / 分组 / 隐藏行列？ | 见 §14.1 |
| Q10 | ~~用户已把目标文件开着~~ → **用户确认：约定范畴，按 SOP 处理** | 已关闭 |

### 14.1 筛选是否允许（回答「是否允许筛选」）

**结论：允许筛选查看，但不影响合并。**

- 合并基于 `Base/Local/Remote` 的**整块列区域**，与"当前可见行"无关
- **代码禁止**使用 `SpecialCells(xlCellTypeVisible)` 或对 `Selection` 操作
- **禁止**用户对 study 列做「只粘贴到可见单元格」（本质是用户自己的操作失误，
  靠 diff 发现：隐藏行也会被比较到）
- 排序**必须禁止**（会让行与 Base 错位）→ 结构守卫会拦下，但更应在保护设置里
  关掉 `AllowSorting`

### 14.2 Unix / SAS 打开方式的差异（回答「unix 与 win 的锁区别」）

**这是一个真实且重要的问题，尚未实测。**

已知事实：
- Windows Excel 打开时创建 `~$` owner 文件，并请求 `deny-write / allow-read` 共享模式
- Linux 上 SAS 读文件不创建 owner 文件，通常也不申请 Windows 写锁；
  Samba 默认 `strict locking = no`，可能**允许两端同时打开**
- 极端情况下：SAS 写、Excel 也写 → 文件损坏

因此：
- 不能只靠 `~$` 文件判断占用（它对非 Excel 进程无约束力）
- 需要对共享文件本身做 **Win32 `CreateFile` 独占探测**（已验证可用）
- **[待测]** 见 §16 T5

### 14.3 网络延迟是否导致读到过期内容（回答）

提交时「打开 → 读 → 写 → 保存」窗口内，读的是 `Workbooks.Open` 的最新内容，
SMB 的 oplock/lease 失效机制会保证读到最新数据。

残余风险：
- 两个用户提交**不同列**时，理论上第二个人的读可能早于第一个人的写完成
  → 但他们的列不重叠，合并结果无冲突，且第二个人写入时只写自己的列
- 真正的隐患是**整文件级**的读-改-写覆盖 → 但我们**只写声明列**，不整文件覆盖，
  所以即使读到旧内容，也不会覆盖他人的列

**结论：列级写入让这个风险可接受，无需额外机制。**

---

## 15. sheet 保护方案

### 15.1 ✅ 已决定：改模板 —— product 列加锁，AMG 列解锁

用户答复（2026-11-18）：**「我们可以默认可以开锁，如果后面需要再加对应的功能」**
→ 采用本方案，**工具完全不需要密码**。原 §15.2 的密码路径降级为"若模板被还原时的兜底"。

保持 sheet 保护开启，但把 **AMG 列的 `Locked` 属性设为 False**。

这样：
- 用户可在受保护的表里正常编辑 study 列（Excel 原生能力）
- 工具写回时也不需要密码（study 列本来就解锁）
- 无需**任何**密码分发，最干净

**代价**：一次性模板改造（每个文件类型改一次）。

### 15.2 若模板不能改：用户输入密码

```
Import：对签出的 sheet → 弹窗输入密码 → Unprotect（密码仅存内存）
        记录原始保护标志位（AllowFormattingCells / AllowInsertingColumns / AllowFiltering ...）
Save  ：Unprotect → 写声明列 → 按原标志位 Protect(同一密码) → Save
```

必须用 `finally` 保证重新保护，否则会以无保护状态存回。

### 15.3 现状

- 42 张 sheet 为 SHA-512 强密码
- 无 GxP 约束 → 程序化解除在合规上可接受
- `.xlsx` 的 sheet 密码**不是加密**（删 XML 元素即可解），只是防误操作护栏

---

## 16. 待测清单（按优先级）

| # | 测试 | 为什么必须 | 阻塞什么 |
|---|---|---|---|
| **T1** | ~~健康环境基线~~ | ✅ **已完成**（测试 7，VBA 内网）：open 3.27 / save 6.77 / 合计 9.41 s | — |
| ~~T21~~ | ~~AMG 大小写撞名检测~~ | ❌ **已作废**（用户确认 P1 业务上不可能）。保留 3 行运行时断言防静默错误（§3.3.1） |
| ~~T22~~ | ~~锚点唯一性实测~~ | ❌ **已作废**（用户确认 P2 是业务语义）。仅保留「跳过空锚点行」与「无 product 列报错」两项机械守卫 |
| **T2** | **`MkDir` 跨机原子性**（两台机器同时抢同一目录） | 锁的全部正确性 | 锁实现 |
| **T3** | **Excel 保存原子性**：Save 过程中杀掉 Excel，原文件是否完好 | 唯一真正的数据风险 | 恢复策略 |
| **T15** | ~~事件屏蔽有效性~~ | 已完成：事件归零；`ThisWorkbook` 事件不被 candidate 触发 | — |
| **T16** | ~~目标文件已被打开~~ | 已完成：返回同一个 Workbook；已归入约定范畴 | — |
| **T17** | `AutomationSecurity = ForceDisable` 下打开 `.xlsm` | 低优先级；本工具只开 `.xlsx` | 可选 |
| **T4** | **`Range.Copy` 后的保护/解锁往返**：受保护 sheet Unprotect→写→Protect 后标志位是否完整还原 | 否则改变用户环境 | 保护方案 A |
| **T5** | **SAS/Linux 并发读写的实际行为** | 可能损坏文件 | 是否有额外防线 |
| **T6** | **崩溃后锁回收**：心跳机制能否在 3 分钟内被他人正确判定为过期 | 「崩溃后不残留」的唯一保障 | 锁实现 |
| **T7** | ~~颜色读取稳定性~~ | ✅ **已完成**：保存/重开无漂移；主题切换会造成漂移（已加主题指纹防护） | — |
| **T8** | ~~diff 耗时~~ | ✅ **已完成**：2,000 格读样式 ≈ 0.08 s，完全可接受 | — |
| **T9** | ~~`Value(11)` 对比~~ | ✅ **已否决**：非确定性，不能做指纹 | — |
| **T10** | **Ribbon 注入 .xlam** 在本 Excel 版本可用 | 否则没有按钮 | 部署形态 |
| **T11** | 只读读者不阻塞写入者 | ✅ **已完成** | — |
| **T12** | `Range.Copy` 保真性 21 项 | ✅ **已完成** | — |
| **T13** | `Open...Lock` / `CreateFile` / `MkDir` 本机互斥 | ✅ **已完成** | — |
| **T14** | FNV-1a 64 位正确性 | ✅ **已对拍** | — |

**T1 已完成（测试 7）。第 12 节数字均为实测，可直接写入用户文档。**

---

## 17. 残余风险

| 风险 | 等级 | 说明 |
|---|---|---|
| 提交窗口不可预测（共享盘阶段性降级） | 高 | 已有预检 + 15 分钟超时 + 进度条；无法根治 |
| Excel 崩溃时锁不释放 | 中 | 靠心跳 + 过期清理，最长残留约 3 分钟 |
| `Save()` 中途失败的原子性未经实测 | 中 | T3 |
| 用户绕过工具直接改文件 | 中 | Base diff 能发现并转成冲突；无法阻止 |
| 结构变更（增删行列） | 中 | 结构守卫会中止提交；需工作组配合先查锁 |
| SAS 与 Excel 并发写 | 低–中 | T5 |
| 时钟偏移导致过期误判 | 低 | 以心跳 mtime 为主判据 |
| 颜色假冲突 | 低 | T7 已测；已加主题指纹防护 |
| 幽灵行拖慢打开/保存 | 中 | 不在范围；建议单独立项清理 |

---

## 18. 一句话

> **签出 = 按列加锁 + 落 Base 快照 → 本地编辑 → 提交时打开当前文件、
> 结构守卫、三方合并、只写声明列、一次 Save、释放锁。**
> 慢的部分（打开/保存）无法优化，只能用预检 + 进度 + 超时来管理用户预期。
