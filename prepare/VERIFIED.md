# 已验证的事实（本机实测结果）

> 环境：Windows 11 Pro / Excel 16.0.19127.20860 **x64** / PowerShell 5.1 + 7.6
> 测试日期：2025-11-16
> 这些是**已确认**的结论，不再需要在你工作电脑上重复验证。

---

## 1. `Range.Copy` 保真性 —— ✅ 全部通过

用 `6-test-copyfidelity.bas` 实测（跨工作簿、含超链接/样式/批注/验证/条件格式）：

| 项目 | 结果 |
|---|---|
| 值 `Value2` | ✅ 保留 |
| **超链接数量** | ✅ 保留（2 → 2） |
| **内部跳转 `SubAddress`** | ✅ 保留（TOC 跳转可用） |
| **外部链接 `Address`** | ✅ 保留 |
| `TextToDisplay` | ✅ 保留 |
| 填充色 `Interior.Color` | ✅ 保留 |
| 字体色 / 粗体 / 斜体 / 字号 / 字体名 | ✅ 全部保留 |
| 数字格式 `NumberFormat` | ✅ 保留 |
| 合并单元格 | ✅ 保留 |
| 批注 | ✅ 保留 |
| 数据验证 | ✅ 保留 |
| 条件格式 | ✅ 保留 |
| 水平/垂直对齐、边框 | ✅ 保留 |
| **行高 / 列宽** | ❌ **不携带**（Excel 固有行为） |
| 冻结窗格 | ❌ 不携带（窗口级设置） |

**结论**：`Range.Copy Destination:=` 可以作为写回手段，超链接不会丢。
**已知限制**：行高/列宽需在 Copy 后**显式复制**（已实测可行，3 行代码）：

```vba
dst.Rows(1).RowHeight        = src.Rows(1).RowHeight
dst.Columns(2).ColumnWidth   = src.Columns(2).ColumnWidth
```

> 你说过列宽不关心，所以这一条影响很小；实现时顺手加上即可。

---

## 2. `Value2` 数组写入会丢失全部样式和超链接 —— ✅ 已确认

实测对照：

```
Value2 写入超链接数 = 0  (源=2)
Value2 写入 A1 填充色 = &HFFFFFF  (源=&HC47244)
```

**这证实了"必须用 Copy 而不是 Value2"这个核心设计决策。**
（`Value2` 只适合用于**哈希计算**，绝不能用于写回。）

---

## 3. 缩小写入必然残留 —— ✅ 已确认，必须 Clear 并集

实测三步：

| 步骤 | 操作 | A5 | A10 |
|---|---|---|---|
| 1 | 目标写入 10 行旧数据 | OLD5 | OLD10 |
| 2 | **只覆盖** A1:B4（4 行新数据） | **残留 OLD5** ❌ | **残留 OLD10** ❌ |
| 3 | 先 Clear A1:B10（并集）再覆盖 A1:B4 | 空 ✅ | 空 ✅ |

**结论**：写回前清理范围必须是 `Union(服务器当前区域, 本地区域)`，
否则用户"删行"会表现为"没删掉"——这是最容易让用户失去信任的 bug。

---

## 4. VBA 的 `Open ... Lock` 互斥 —— ✅ 语义正确

`4-vba-lock.bas` 自测结果：

```
1) 首次独占: ✅ 成功 (句柄 1)
2) 重复独占: ✅ 被拒绝 70 (Permission denied)
3) Lock Write 打开(读): 被拒绝 70（符合预期）
4) 释放后独占: ✅ 成功
5) Name 覆盖本地文件: ❌ 失败 58 File already exists
6) 100 次 加锁+解锁 (本地盘): 23 ms   → 约 0.23 ms/次
```

**结论**：互斥语义正确；加锁开销约 0.2ms，可忽略。

### ⚠ 重要发现：`5) Name 语句不能覆盖已存在文件`

VBA 的 `Name src As dst` 在 `dst` **已存在**时报 `error 58`。
这意味着**不能用 `Name` 实现"重命名换页"**。

正确做法（三选一，按推荐度）：

| 方案 | 写法 | 说明 |
|---|---|---|
| **A** | `Declare PtrSafe Function MoveFileExW` + `MOVEFILE_REPLACE_EXISTING` | 真覆盖，与 `5-test-replace.ps1` 测的同一 API |
| B | `Object.MoveFile`（FileSystemObject） | 有 `OverWriteFiles:=True` 参数 |
| C | `Kill dst` 然后 `Name src As dst` | **不原子**，中间有窗口，不推荐 |

→ 若 `5-test-replace.ps1` 证明 Samba 允许覆盖被打开文件，则用方案 A。

---

## 5. Win32 `CreateFile` 独占互斥 —— ✅ 本机生效

`4-test-lock.ps1` 的 C# 编译并运行通过：

```
CreateNewExclusive            -> valid=True
OpenExclusive(无人持有)        -> valid=True
OpenExclusive(已持有)          -> valid=False  err=32 (SHARING_VIOLATION)
```

`mkdir` 原子锁同样生效：

```
A: CreateDirectory(.lock_xxx_HELD)
B: CreateDirectory(同名) -> False, ERROR_ALREADY_EXISTS(183)
```

**两种备用锁原语都可用**（SMB 上的真实行为仍需两机实测）。

---

## 6. `MoveFileEx(WITH REPLACE)` 语义 —— ✅ 本机正确

```
正常替换:        True    err=SUCCESS(0)
目标被独占打开:   False   err=ACCESS_DENIED(5)
```

**注意**：Windows 本地盘返回的是 **5 (ACCESS_DENIED)**，不是 32。
Samba 上可能返回 32。判定时要**同时接受 5 和 32**，并结合"释放后重试成功"来确认。
（`5-test-replace.ps1` 已加入这个判定说明。）

---

## 7. FNV-1a 哈希在 VBA 里的正确写法 —— ✅ 已对拍验证

**踩到的坑（正式实现必须避开）**：

- `h * 16777619` 用 `Long` 会溢出
- 用 `Mod` 运算符**也不行** —— VBA 的 `Mod` 会把操作数强制转 `Long`，而乘积约 7.2e16
- 用 `Double` 直接乘会丢精度（尾数只有 53 位）

**可用算法**（`Double` 分段精确计算，已与标准 FNV-1a 对拍一致）：

```vba
Private Function FnvStep(ByVal h As Long, ByVal c As Long) As Long
    Const P As Long = 16777619
    Dim u As Double, hi As Double, lo As Double, r As Double
    h = h Xor (c And &HFF&)
    If h < 0 Then
        u = CDbl(h) + 4294967296#
    Else
        u = CDbl(h)
    End If
    hi = Int(u / 65536#)
    lo = u - hi * 65536#
    r = (hi * P - Int(hi * P / 65536#) * 65536#) * 65536#
    r = r + (lo * P - Int(lo * P / 4294967296#) * 4294967296#)
    r = r - Int(r / 4294967296#) * 4294967296#
    If r > 2147483647# Then r = r - 4294967296#
    FnvStep = CLng(r)
End Function
```

实测对拍（VBA 输出 vs 标准 FNV-1a 32bit）：

```
"a"             E40C292C  = 0xE40C292C  ✅
"abc"           1A47E90B  = 0x1A47E90B  ✅
"hello"         4F9F2CAB  = 0x4F9F2CAB  ✅
"Sheet3"        421D4FA1  = 0x421D4FA1  ✅
"K001|Alice|100" 8028AD82 = 0x8028AD82  ✅
```

双线（h1/h2 错开）拼接成 64 位，用于区域哈希。

---

## 8. 前导零 key 的规则 —— ✅ 已确认机理

测试中 `"00123"` 一开始变成 `123`，原因：**赋值时该单元格还是"常规"格式**。

**规则（必须写进使用协议）**：

> key 列的文本格式 `NumberFormat = "@"` 必须在**写入值之前**设置。
> 若已有历史数据是数字格式，需先做一次性的格式修正，否则 Excel 会在用户编辑时把 `00123` 变成 `123`，导致 key 漂移。

验证过：先设 `"@"` 再赋值，`"00123"` 完整保留。

---

## 9. Excel 版本信息（本机，非工作机）

```
Excel 16.0.19127.20860  x64  (ClickToRun: ProfessionalRetail, ProPlus2019Volume)
OS: Windows 11 Pro
```

工作机的需由 `1-office-env.ps1` 采集。

---

## 10. 仍需在工作机验证的（已收缩）

| # | 项目 | 脚本 | 为什么本机测不了 |
|---|---|---|---|
| 1 | 宏策略 / 受信任位置 / 网络位置 | `1-office-env.ps1` | 域策略不同（✅ 已完成） |
| 2 | profile 持久性 | `1-office-env.ps1` ×2 | ✅ 已完成 |
| 3 | 真实文件的 sheet/样式/key 画像 | `2-file-profile.ps1` | ✅ 已完成 |
| 4 | 网络耗时、Excel Save 到 Samba | `3-measure.ps1` | ✅ 已完成 |
| 5 | **SMB 上的锁是否真互斥**（两机） | `4-test-lock.ps1` A/B | ❌ **仍待测（T2）** —— 需两台机器 |
| 6 | Samba 是否允许覆盖被打开文件 | `5-test-replace.ps1` | ❌ 仍待测 |
| 7 | 真实文件的 Copy 保真性 | `6-test-copyfidelity.bas` | ✅ 21/21 PASS |
| 8 | Excel 实例策略 / 内网基线 | `测试结果.xlsx`（VBA） | ✅ 已完成（测试 7） |
| 9 | **Save 原子性**（Save 中杀 Excel） | 未写 | ❌ **T3 仍待测** |

---

## 12. T8 / T9 / T15 / T16 实测结果（开发机，Excel 16.0 build 19127，x64）

> 脚本：`collect/8-test-diffperf.bas`、`8b-test-diffedge.bas`、`8c-test-difffinal.bas`
> 驱动器：`_smoke/run_diffperf.py`（自动建 .xlsm、导入模块、编译探针、跑测试）
> 计时：`QueryPerformanceCounter`，所有临时工作簿均隐藏。

### 12.1 T15：`EnableEvents = False` 真的能压住目标工作簿的 `Workbook_Open`

| 条件 | marker 文件是否产生 | 期望 |
|---|---|---|
| `EnableEvents = False` | **False** | False ✅ |
| `EnableEvents = True` | **True** | True ✅ |

并在 `Workbooks.Open` 之后回读 Application 状态：

```
EnableEvents=False  ScreenUpdating=False  Calculation=-4135
```

→ **状态没有被目标工作簿偷改。**

**结论**：改用宿主实例（不新建 Excel）安全，前提是：
1. 进入时 `EnableEvents = False`
2. 自己模块里所有事件处理第一行 `If mInCommit Then Exit Sub`
3. `finally` 恢复原状态

### 12.2 T16：目标文件已经开着时的行为

| 场景 | 结果 |
|---|---|
| 对已 r/w 打开的文件再次 `Workbooks.Open(p, 0, False)` | **0.006 s 返回同一个 Workbook 对象**（`wb1 Is wb2 = True`） |
| 再以 read-only 打开同一个文件 | 也是**同一个对象**，不新建窗口 |
| `Windows.Count` | 始终 1 |

→ Excel **不会**报错、**不会**建第二个副本。

**结论**：提交前如果发现目标文件已经被用户开着，
可以直接复用那个 Workbook（但必须先确认 `Saved = True`，否则用户有未保存修改），
或者提示用户关闭。**绝对不能直接 `Save`，否则会把用户未保存的修改一起提交掉。**

### 12.3 T8：读取与 diff 耗时（关键数据）

| 块 | 读 `.Formula` | 读 `.Value2` | 纯 VBA diff | 哈希数组 | `Range.Copy` |
|---|---|---|---|---|---|
| 500×10 = 5k | 0.004 s | 0.000 s | 0.002 s | 0.027 s | 0.025 s |
| 2000×30 = 60k | 0.053 s | 0.007 s | 0.021 s | 0.402 s | 0.088 s |
| **5000×30 = 150k** | **0.143 s** | **0.033 s** | **0.062 s** | **0.932 s** | **0.226 s** |

### 12.4 T8c：非连续列的读取（真实形状）

真实场景下 study 列是**分散**的，不是连续块。

| 方式 | 耗时（150k 格） |
|---|---|
| 30 列逐列 `Range(...).Formula` 读，再在 VBA 里拼数组 | **0.152 s** |
| 读包围的连续块（5000×90 = 450k 格）再挑 | **0.453 s** |

→ **逐列读更快**（0.152 vs 0.453），因为只读需要的列。
→ 但 `Union` 不能用，见 12.5。

### 12.5 ⚠️ 重大陷阱 1：`Union` 的多区域读取会**静默丢列**

```
Union 6 列 → Areas.Count = 6
Union.Cells.Count = 2000      ← 实际有 12,000 个单元格！
Union.Formula → UBound1=2000, UBound2=1   ← 只有 1 列！
Union.Value2  → UBound1=2000, UBound2=1   ← 只有 1 列！
```

`Union.Formula` **只返回第一个 Area**，没有任何报错。

同样地，`Union.Copy Destination:=` **只复制第一个 Area**：

```
multi-area Copy -> B1 : 0.017 s
  dst B1=B1  E1=<空>  G1=<空>     ← E、G 两列完全没写过去
```

→ **如果写回时用 `Union(study 列).Copy`，会静默丢掉除第一列外的所有列。**
→ 这是**数据丢失级**的坑，绝对不能这么写。

### 12.6 ⚠️ 重大陷阱 2：混合区域读属性返回**不可靠值**

对 `A1:A2`（A1 红/粗/0.00/左对齐/有边框，A2 蓝/普通/General/右对齐/无边框）：

| 属性 | Null? | 返回值 | 可信？ |
|---|---|---|---|
| `Interior.ColorIndex` | True | NULL | ✅ 正确（表示"不一致"） |
| **`Font.Bold`** | True | NULL | ✅ |
| **`Font.Color`** | True | NULL | ✅ |
| **`NumberFormat`** | True | NULL | ✅ |
| **`HorizontalAlignment`** | True | NULL | ✅ |
| **`Interior.Color`** | **False** | **0** | ❌ **错！A1 实际是 255** |
| **`Borders(7).LineStyle`** | **False** | **1** | ❌ **只看第一个单元格** |

→ `Interior.Color` 和 `Borders().LineStyle` 在混合区域**不返回 Null，而是静默返回单格值**。
→ **绝对不能靠 "Null 则回退逐格读" 这个逻辑**，因为 `Interior.Color` 永远不返回 Null。

**另一个坑**：`RGB(9,9,9)` 读出 `592137`，而 `9 + 9*256 + 9*65536 = 591366`。
→ 差 771。`Interior.Color` 返回的是 **BGR 且含保留位**，**不要自己拼 RGB 比对**，
   应该一律用 `Interior.Color` 原值比较，不要拆字节。

### 12.7 T9：`Value(11)`（XML Spreadsheet）是否可用

**能捕获的内容**（`A1:C1`，1286 字符）：

| 属性 | 是否包含 |
|---|---|
| Bold / Italic | ✅ |
| Font（名称、字号、颜色） | ✅ |
| Interior（填充） | ✅ |
| NumberFormat | ✅ |
| Border | ✅ |
| Alignment | ✅ |
| **MergeAcross（合并格）** | ✅ |
| **HRef（超链接）** | ✅ |
| **Comment（批注，含正文）** | ✅ |

**速度**：

| 块 | XML 长度 | 读取耗时 |
|---|---|---|
| 5,000 格 | 0.35 M 字符 | 0.043 s |
| 60,000 格 | 4.23 M 字符 | 0.187 s |
| 150,000 格 | — | 跳过（XML 过大） |

### 12.8 ⚠️ 重大陷阱 3：`Value(11)` **不是确定性的**

同一区域连续读两次：**相同**。
但 **保存并重开后，以及颜色改了再改回来之后，都不相同**。

差异位置固定在第 591 字符附近，内容形如：

```
A: [63">  <Font ss:FontName="等线" ss:Size="11" ss:Color="#000000" ss:Bold="1"/>]
B: [16">  <Font ss:FontName="等线" ss:Size="11" ss:Color="#000000" ss:Bold="1"/>]
```

→ `Value(11)` 内部含**每次生成都不同的 StyleID 计数器**（`63` vs `16`）。
→ **绝对不能拿 `Value(11)` 的哈希当指纹/用于判定"有没有变"**，
   否则每次都会报"全部变了"，产生海量假冲突。

### 12.9 T9b：字符串比较 vs 哈希（性能取舍）

对 4.2 M 字符的字符串：

| 操作 | 耗时 |
|---|---|
| `StrComp`（相同） | **0.002 s** |
| `StrComp`（首字节不同） | **0.000 s** |
| 哈希前 200k 字符 → 外推全程 | **1.527 s** |

→ **`StrComp` 快 760 倍。**
→ VBA 的 `StrComp` 是原生 C 实现，哈希是逐字符 VBA 循环，后者慢得多。

**结论**：
- 小范围/需存指纹 → 用哈希
- **大块内容比较 → 直接 `StrComp`，不要自己哈希**

### 12.10 T7：颜色稳定性的实测（决定会不会有假冲突）

脚本：`collect/8d-test-theme.bas`（`ZZRUNTHEME`，报告 `%TEMP%\mddt_theme.txt`）

#### T7a 主题色与显式 RGB 的等价性

| 设置方式 | `.Interior.Color` | `.ThemeColor` | `.TintAndShade` |
|---|---|---|---|
| `ThemeColor=5, Tint=0` | **12874308** | 5 | 0 |
| 同值显式 `RGB` | **12874308** | 0 | 0 |
| `ThemeColor=5, Tint=0.4` | **14395790** | 5 | 0.39998 |
| 同值显式 `RGB` | **14395790** | 0 | 0 |

→ `.Color` **会把主题色解析成 RGB**，所以「主题色」与「同值显式色」的 `.Color` 相同。
→ 差异只在 `ThemeColor` / `TintAndShade`。
→ **只比 `.Color` 即可判定"看起来是否一样"，这正是变更判定需要的语义。**

注意 `RGB(r,g,b)` 在 VBA 里是 `r + g*256 + b*65536`（BGR 顺序）。

#### T7e 「无填充」与「白色填充」无法区分（但无害）

| 状态 | `.Color` | `.ColorIndex` | `.Pattern` |
|---|---|---|---|
| 完全未触碰的单元格 | **16777215** | -4142 | -4142 |
| `Pattern = xlNone` | 16777215 | -4142 | -4142 |
| 显式白色 | **16777215** | **2** | 1 (Solid) |
| `Color = 16777215` | 16777215 | 2 | 1 |

→ 默认值与"无填充"都是 `16777215`，与显式白色无法用 `.Color` 区分。
→ **视觉上完全一致，对本工具无害**；且我们写回用 `Range.Copy`（连着 `Pattern` 一起复制），
   不会把"无填充"变成"实心白"。

#### T7f 字体色

| 状态 | `.Font.Color` | `.Font.ColorIndex` |
|---|---|---|
| 默认 | **0** | 1 |
| `RGB(255,0,0)` | 255 | 3 |
| `ThemeColor=5` | 12874308 | — |
| `ColorIndex = xlAutomatic` | **0** | -4105 |

→ 默认与"自动"都是 `0`，视觉一致，无害。
→ 字体色同样可用 `.Color` 单值比较。

#### ⚠️ T7b 主题切换会改变已有的主题色填充（唯一真风险）

| | 切换前 | 切换后 | 变了吗 |
|---|---|---|---|
| 主题色填充 | 12874308 | **2540176** | **是** |
| 显式 RGB 填充 | 3678732 | 3678732 | 否 |

更糟：**把主题换回去之后，值也没有恢复**（变成 8544277，不等于原值 12874308）。

→ 若有人改了工作簿主题，会凭空产生"颜色全变了"的假冲突。

#### ✅ T7c 保存 + 重开：完全不变（最重要的一条）

| 项 | 保存前 | 重开后 | 一致 |
|---|---|---|---|
| 填充色 | 3302600 | 3302600 | ✅ |
| 字体色 | 15759370 | 15759370 | ✅ |
| `ColorIndex` | 46 | 46 | ✅ |
| 主题色填充 | 14395790 | 14395790 | ✅ |

→ **正常"编辑→保存→重开"路径不会产生任何颜色漂移。**
→ 这是本工具的主路径，因此**颜色比较可以安全启用**。

#### ⚠️ T7d 跨主题工作簿 Copy

| | 源 | 目标（不同主题） | 一致 |
|---|---|---|---|
| 显式 RGB 填充 | 3302600 | 3302600 | ✅ |
| 字体色 | 15759370 | 15759370 | ✅ |
| **主题色填充** | 14791492 | **7594943** | ❌ |

`Range.Copy` **连主题引用一起复制**（目标 `ThemeColor=5, Tint=0.4` 保留），
因此在主题不同的工作簿里解析出不同的 RGB。

**但本工具不受影响**：本地副本是用 `SaveCopyAs` 从服务器文件派生的，
**继承同一个主题**。只有当用户从别的工作簿粘贴带主题色的内容时才会出现。

#### 结论与设计决定

| 决定 | 依据 |
|---|---|
| 颜色比较用 `.Interior.Color` / `.Font.Color` 单值 | T7a / T7f：已归一化为 RGB，语义正确 |
| **不**比较 `ThemeColor` / `TintAndShade` | 会让"同一颜色不同来源"变成假冲突 |
| 颜色比较**默认可开启** | T7c：保存/重开无漂移 |
| 增加**主题指纹**校验 | T7b：主题变更才会造成假冲突，需能检测 |
| 主题指纹做法 | 临时单元格设 `ThemeColor=1..12, Tint=0` 各读一次 `.Color`，拼成 12 个值做 FNV 哈希（≈12 次属性写+读，毫秒级）；Import 时存进 meta，Commit 时比对，不一致则告警 |

### 12.11 汇总：最终实现选型

| 需求 | 采用 | 放弃 | 理由 |
|---|---|---|---|
| 读值 | 逐列 `Range().Formula` 后拼数组 | `Union.Formula` | 12.5：Union 丢列 |
| 读样式 | 逐格 `.Interior.Color` / `.Font.Color` / `.Font.Bold` | 整块属性读 | 12.6：混合区域静默返回错值 |
| 样式指纹 | **不用 `Value(11)`** | `Value(11)` 哈希 | 12.8：非确定性 |
| 大块比较 | `StrComp` | 自实现哈希 | 12.9：慢 760 倍 |
| 指纹存储 | FNV-1a 64（仅用于跳过细比） | MD5 | 已对拍，够用 |
| 写回 | 逐列 `Range.Copy Destination:=` | `Union.Copy` | 12.5：Union 只写第一列 |
| 大块写入 | 逐列 Copy | `Value2` 数组 | 早期实测：Value2 丢样式与超链接 |

---

## 13. 宿主实例方案：事件与焦点隔离（已全量实测）

> 脚本：`_smoke/run_hostevents2.py`（构建**真实 .xlam**，`IsAddin=True`，在插件内部
> 挂 `Application` 级事件，再让插件自己去打开/关闭一个 candidate workbook）
> 日志：`%TEMP%\mddt_hostev2.log`
>
> 前提：Excel 16.0 build 19127 x64。测试用 `.xlsx` candidate，非宏文件。

### 13.1 插件自身的 ThisWorkbook 事件：**不会被 candidate 触发**

加载 .xlam 时的完整日志：

```
PHASE          | before opening the .xlam
THISWB:Workbook_Open | mddt2.xlam
PHASE          | after opening the .xlam
```

**只有 `ThisWorkbook.Workbook_Open` 触发**，且只在**插件自己被加载时**触发一次。

标准模块里故意放的两个入口：

```vba
Public Sub Workbook_Open()   ' 标准模块里的 Workbook_Open
Public Sub Auto_Open()       ' 标准模块里的 Auto_Open
```

**两者都没有被调用。** close 时同理：`Auto_Close` 也没有触发，
只有 `THISWB:Workbook_BeforeClose`。

然后在 candidate 的 open/close 全过程（Q2/Q3）里，
**没有再出现任何 `THISWB:` 事件**。

→ **答案：打开/关闭 candidate 不会触发插件自己的 `ThisWorkbook.Workbook_Open`
   或 `Workbook_BeforeClose`。**
→ `BeforeClose` 误释放锁的担心 **不成立**（前提：锁释放逻辑挂在
   `ThisWorkbook`，不要挂 `Application` 级）。
→ **推论：绝不要把业务逻辑写进标准模块的 `Workbook_Open` / `Auto_Open`** ——
   实测不会被调用，会静默失效。

### 13.2 candidate open/close 到底触发哪些事件（`EnableEvents = True`）

| 阶段 | 触发的事件（按顺序） | 参数 |
|---|---|---|
| `Workbooks.Open` | `App.WorkbookOpen` | candidate |
| | `App.WorkbookActivate` | candidate |
| | `App.WindowActivate` | candidate |
| `wb.Close(False)` | `App.WorkbookBeforeClose` | candidate |
| | `App.WindowDeactivate` | candidate |
| | `App.WorkbookDeactivate` | candidate |

**共 3 + 3 个，全部是 `Application` 级事件，无一属于 `ThisWorkbook`。**
（`SheetActivate` / `SheetSelectionChange` 未触发，因为过程中没有激活 sheet 或改动选区。）

### 13.3 `EnableEvents = False` 后：**事件归零**

同一套 open/close 循环，`EnableEvents = False`：

```
Q3  | same cycle with EnableEvents=FALSE
Q3  | OPEN returned
Q3  | CLOSE returned
```

**日志里 0019–0024 之间没有任何 `APP:` 事件。**

→ **完全抑制，不是"部分抑制"。**

### 13.4 ✅ 无感后台处理：**可以做到**（关键证据）

`ScreenUpdating = False` + `EnableEvents = False`，Open 后**立刻隐藏窗口**：

| 时刻 | 活动工作簿 | 活动表 | 选区 |
|---|---|---|---|
| `SNAP:BEFORE`（用户态） | Book3 | Sheet1 | **D5** |
| `SNAP:MID-open`（刚 Open 完） | **candidate** ← 被抢 | AE | A1 |
| **`SNAP:MID-hidden`（Visible=False 之后）** | **Book3** | **Sheet1** | **D5** ← 立刻回来 |
| `SNAP:AFTER-close` | Book3 | Sheet1 | D5 |
| `SNAP:RESTORED` | Book3 | Sheet1 | D5 |

→ **焦点/活动表/选区只在 `Open` 到 `Visible = False` 之间的几毫秒内被抢，
   一旦隐藏窗口就完整归还，且用户的选区（D5）没有丢失。**
→ 配合 `ScreenUpdating = False`，用户**看不到任何重绘**。
→ 暴露窗口 = "几毫秒 + 无重绘"，且此时用户正被我们的模态进度窗体挡住。

**隐藏后所有操作照常可用**（round 1 的 Scenario C 实测）：

| 操作 | 隐藏工作簿上的结果 |
|---|---|
| `Range.Copy Destination:=`（隐藏→隐藏） | **OK**，值 `LOCAL1/LOCAL2`、填充色 `16764057` 均正确 |
| `wb.Save` | **OK** |
| `Range.Clear` | **OK** |

→ **整个提交流程可以在一个不可见的 workbook 上完成。**

### 13.5 推荐做法（已验证）

```vba
Private mInCommit As Boolean

Sub Commit()
    Dim oldEE As Boolean, oldSU As Boolean
    oldEE = Application.EnableEvents
    oldSU = Application.ScreenUpdating

    mInCommit = True                    ' 第二道防线
    On Error GoTo Cleanup

    Application.EnableEvents = False    ' 事件归零（13.3 实测）
    Application.ScreenUpdating = False  ' 不重绘

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(remotePath, 0, False)
    wb.Windows(1).Visible = False       ' 焦点立刻归还（13.4 实测）

    ' ... 结构守卫 / 三方合并 / 逐列 Copy / wb.Save ...

    wb.Close False
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = oldSU
    Application.EnableEvents = oldEE
    mInCommit = False
End Sub

' 所有 Application 级事件处理的第一行：
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    If mInCommit Then Exit Sub
    ...
End Sub
```

### 13.6 ❌ `Application.Interactive = False` **不需要，且不建议**

| 观察 | 结果 |
|---|---|
| `Interactive = False` 时新打开的 workbook 窗口是否自动隐藏 | **否**，仍然 `Visible = True` |
| 加上 `Interactive = False` 后焦点归还行为 | 与不加完全一样（SNAP 序列逐项相同） |

→ 它**不能**解决隐藏问题，实测无额外收益。
→ 若在 `Interactive = False` 状态下异常退出，用户整个 Excel 会**失去响应**
  → 这是一个比"焦点闪一下"严重得多的事故模式。
→ **决定：不使用。**

### 13.7 剩余的真实风险（更新后）

| 风险 | 等级 | 结论 |
|---|---|---|
| 抢焦点 / 界面乱跳 | **已消除** | 13.4：Open 后立即 `Visible=False`，选区完整归还 |
| 触发插件自身 `Workbook_Open`/`BeforeClose` | **不存在** | 13.1：ThisWorkbook 事件不由 candidate 触发 |
| 触发 `Application` 级事件导致递归 | **已消除** | 13.3：`EnableEvents=False` 事件归零；再叠 `mInCommit` |
| 提交中途异常 → 全局 `EnableEvents` 卡在 False | **中** | 必须 `On Error GoTo Cleanup` + Cleanup 内置 `Resume Next` |
| candidate 是 `.xlsm` 且带宏 | 低（本工具只开 `.xlsx`） | 建议加 `AutomationSecurity = msoAutomationSecurityForceDisable`（**未测，T17**） |
| 用户已把目标文件开着 | **约定范畴** | 用户确认：按约定处理，工具不为此设计额外流程 |

## 14. Workbook 角色识别与角色化事件处理（已实测）

> 脚本：`_smoke/run_roles3.py`（真实 `.xlam`，`IsAddin=True`）
> 日志：`%TEMP%\mddt_roles3.log`
> 目标：能否严格识别「这个 workbook 是什么角色」，并让 `Application` 级
> 事件处理器**按角色分流**。

### 14.1 R1：在事件处理器里**无法找到**插件自身

```
R1 | workbooks=0
R1 | matches found = 0
```

`ThisWorkbook.IsAddin = True`，但 **`Application.Workbooks` 里根本没有它** —— count = 0。

→ 插件工作簿**不在 `Workbooks` 集合中**（Excel 的 add-in 不占这个集合）。
→ 遍历 `Workbooks` 找自身**是错的**。
→ 在处理器里判断"这是不是我"只能靠 `Wb Is ThisWorkbook`。

### 14.2 R2/R3：挂到单个 Workbook 的 `WithEvents` 是**严格的按角色隔离**

```
>>> write to mLocal  →  WB[LOCAL]:SheetChange   + APP:SheetChange
>>> write to mOther  →  WB[OTHER]:SheetChange   + APP:SheetChange
>>> activate mOther  →  （无 WB[LOCAL] 事件）
>>> activate mLocal  →  WB[LOCAL]:Activate
```

**每个 workbook 只触发自己那一份**，互不串扰。

`Wb Is mLocal` / `Wb Is mOther` 判定全部正确（日志里 `[IsLocal=TRUE]` / `[IsOther=TRUE]`）。

### 14.3 ⚠️ R4：已关闭 workbook 的引用是**失效引用**，但不会报错 —— 更危险

```
before close: gStale.Name = mddt_r3_stale.xlsx
wb.Close
after close: gStale Is Nothing  -> False     ← 仍"不为 Nothing"
after close: gStale.Name        -> ERROR -2147221080 （自动化错误）
gStale Is mOther                -> False     ← 不抛错，但结果无意义
mOther Is gStale                -> False
Set gStale = Nothing            -> OK
```

**这是最重要的发现**：

| 行为 | 结果 |
|---|---|
| 关闭后 `Is Nothing` | **仍返回 False**（看起来"还活着"） |
| 关闭后访问 `.Name` | **抛错 -2147221080** |
| 关闭后 `Is` 比较 | **不抛错**，返回 False |
| 赋值 `Set x = Nothing` | OK |

→ **`Is Nothing` 完全不能用来判断"workbook 是否还开着"。**
→ 误判后果：以为 `mLocal` 还在，`Is` 比较永远 False → 所有事件都当成"别人的"处理。
→ 更糟：这类错误**不会抛异常**，会静默地把逻辑走错分支。

**正确的存活判定**：遍历 `Application.Workbooks`，比较 `FullName`（大小写不敏感）：

```vba
Function IsAlive(ByVal wb As Workbook) As Boolean
    Dim i As Long
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Dim target As String
    target = wb.FullName
    If Err.Number <> 0 Then Err.Clear: Exit Function   ' 失效引用
    On Error GoTo 0
    For i = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(i).FullName, target, vbTextCompare) = 0 Then
            IsAlive = True
            Exit Function
        End If
    Next i
End Function
```

**推论（对设计的直接影响）**：

> **不要依赖长期持有的 workbook 对象引用。**
> 提交流程内用局部变量 + `On Error` 包裹是安全的（生命周期短）；
> 但**跨会话记住"我的本地副本"必须靠路径字符串，不能靠对象引用。**

### 14.4 R5：自定义文档属性可以作角色标记（但**不能用于判脏**）

| 问题 | 实测 |
|---|---|
| `CustomDocumentProperties.Add` 成功？ | **OK**（`MDDT_ROLE = "LOCAL"`，type 4 = 字符串） |
| 加了之后 `wb.Saved` 变 False？ | **不变**（前后都是 False，因为这个 workbook 本来就脏） |
| 经 `SaveAs` 后还在？ | **在**，读回 `LOCAL` |
| 经 `SaveCopyAs` 后副本里有？ | **在**，读回 `LOCAL` |
| 全新 workbook 里有？ | **没有**（`<absent>`） |

→ **可以作为可靠的"角色/来源"标记**，且能跟着副本走。

⚠️ **不能反过来说"属性存在 ⇒ 有未保存修改"**。它对 `Saved` 的影响未单独隔离
（测试时 workbook 本来就是脏的）。**`Saved` 的判定必须另测（T18）。**
但"加属性会把 workbook 弄脏"这个方向是安全的：**属性不会被当成内容丢失**，
`Close(False)` 丢弃它也不会影响原有数据。

### 14.5 R6：挂到 workbook 的类模块**能**取消关闭

```
WB[R6]:BeforeClose | R6
WB[R6]:BeforeClose | Cancel set to True
R6 | after attempted close: workbooks count = 3
R6 | is wb4 still open? True
```

→ `Cancel = True` **确实阻止了关闭**（workbook 仍在）。

⚠️ 同时注意：`APP:BeforeClose` 仍然触发了。说明 **`Application` 级和 workbook 级
是两个独立的订阅**，一个 `Cancel` 不会阻止另一个被调用。

→ 若两处都写"提交中禁止关闭"的逻辑，会**重复提示**。应只在一处实现。

**另外**：从日志时间线可见，`Workbooks.Add` / `Open` 会**自动激活**新工作簿
（`APP:Activate|Book2` 等），即使 `ScreenUpdating = False`。
→ 视觉上不可见，但事件会照常触发 → 再次证明 `EnableEvents` 才是必须的闸门。

### 14.6 由此得到的「角色 + 状态」双轨方案

**结论：你的思路正确，但实现要点与直觉相反 —— 不要用"一个 Application 处理器 + if 分流"，
而要用"Application 级做粗筛 + workbook 级做精确处理"。**

| 层 | 机制 | 用途 | 实测依据 |
|---|---|---|---|
| **L0 闸门** | `Application.EnableEvents = False` + `mInCommit` | 提交期间**所有**事件归零 | §13.3 |
| **L1 粗筛** | `Application` 级 `WithEvents`，处理器第一行按**角色**分流 | 统一的入口、错误包装、状态推进 | §13.2 |
| **L2 精确** | 给**本地副本**单独挂一个 `WithEvents` 类实例 | 只收本地副本的事件，代码里无需再判断是谁 | §14.2 |
| **L3 身份** | 路径字符串（`FullName`）+ `CustomDocumentProperty` | 跨会话识别"哪个是本地副本" | §14.3 / §14.4 |

**角色枚举**（建议）

```vba
Public Enum EWB
    WBR_Unknown = 0      ' 别人的
    WBR_Addin = 1        ' ThisWorkbook（不可达，见 14.1）
    WBR_LocalCopy = 2    ' 我们签出的本地副本
    WBR_RemoteTarget = 3 ' 提交时从服务器打开的
    WBR_UserOther = 4    ' 用户自己的其他工作簿（绝不干预）
End Enum
```

**识别函数（不持有对象引用）**

```vba
Function RoleOf(ByVal wb As Workbook) As EWB
    Dim p As String
    On Error Resume Next
    p = wb.FullName
    If Err.Number <> 0 Then Err.Clear: RoleOf = WBR_Unknown: Exit Function
    On Error GoTo 0

    If StrComp(p, mSession.LocalPath, vbTextCompare) = 0 Then
        RoleOf = WBR_LocalCopy: Exit Function
    End If
    If StrComp(p, mSession.RemotePath, vbTextCompare) = 0 Then
        RoleOf = WBR_RemoteTarget: Exit Function
    End If
    RoleOf = WBR_UserOther
End Function
```

### 14.7 插件级状态机（建议形态）

```vba
' MdlSession
Public Enum ESess
    SS_Idle = 0
    SS_TargetSelected = 1
    SS_CheckingOut = 2
    SS_CheckedOut = 3
    SS_Committing = 4
    SS_Conflict = 5
    SS_Aborting = 6
End Enum

' 处理器统一形状
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error GoTo H
    If mInCommit Then Exit Sub              ' L0 闸门
    If gSess.State = SS_Committing Then     ' 插件级状态
        Cancel = True
        LogWarn "提交期间禁止关闭: " & Wb.Name
        Exit Sub
    End If
    Select Case RoleOf(Wb)                  ' L1 角色分流
        Case WBR_LocalCopy
            HandleLocalClose Wb, Cancel
        Case WBR_RemoteTarget
            ' 不应该发生（提交时已 EnableEvents=False）
        Case Else
            ' 用户的其它工作簿：绝不干预
    End Select
    Exit Sub
H:
    LogError "App_WorkbookBeforeClose", Err.Number, Err.Description
End Sub
```

**必须遵守的规则（全部来自实测）**：

| # | 规则 | 依据 |
|---|---|---|
| 1 | 每个处理器第一行 `On Error GoTo` | 处理器抛错会中断用户操作 |
| 2 | 第二行 `If mInCommit Then Exit Sub` | §13.3 |
| 3 | 只对**确认属于我们**的角色 `Cancel = True` | §14.5：Cancel 很有效，用错会把用户文件锁死 |
| 4 | 不遍历 `Workbooks` 找插件自身 | §14.1：它不在集合里 |
| 5 | 不靠 `Is Nothing` 判断存活 | §14.3：关闭后仍返回 False |
| 6 | 跨会话只用路径字符串识别 | §14.3：对象引用会失效且静默错判 |
| 7 | 关闭拦截只在一处实现 | §14.5：Application 级与 workbook 级是独立订阅，会重复提示 |

## 15. T18：文档属性的「弄脏」行为（已实测）

> 脚本：`collect/8e-test-propdirty.bas`（`ZZRUNPROP`，报告 `%TEMP%\mddt_propdirty.txt`）
> 目的：Import 时想给本地副本打角色标记 `MDDT_ROLE`，需要知道这会不会
> 让 workbook 看起来"有未保存修改"。

### 15.1 结论表

| 操作 | 在干净 workbook 上是否弄脏 | 结论 |
|---|---|---|
| `CustomDocumentProperties.Add` | **是** | ⚠️ 会弄脏 |
| 改已有属性的 `.Value` | **是** | ⚠️ 会弄脏 |
| `.Delete` 删除属性 | **是** | ⚠️ 会弄脏 |
| 读取属性值（200 次） | **否** | ✅ 完全惰性 |
| 读 `Count` | **否** | ✅ 惰性 |
| `BuiltinDocumentProperties("Comments").Value = ...` | **是** | ⚠️ 会弄脏 |

### 15.2 关键数据

```
T18a  Add on clean wb
      Saved after SaveAs   : True    (干净)
      Saved after Add      : False   >>> ADD MAKES IT DIRTY? True

T18b  change existing value
      clean after save     : True
      Saved after change   : False   >>> CHANGE MAKES IT DIRTY? True

T18c  200 reads
      clean before reads   : True
      Saved after 200 reads: True    >>> READ MAKES IT DIRTY? False
      Saved after Count    : True    >>> 惰性

T18h  builtin Comments
      Saved after set      : False   >>> BUILTIN MAKES IT DIRTY? True
```

### 15.3 不影响文件的其他事实

| 事实 | 实测 |
|---|---|
| `Add` 是否立即写盘 | **否**。前后文件大小 8661→8661、mtime 完全不变 |
| 经 `SaveAs` 后属性还在 | **在**（`LOCAL`），且 `Saved = True` |
| 经 `SaveCopyAs` 副本里有 | **在**（`LOCAL`） |
| 重开后还在 | **在** |
| 只读打开（`ReadOnly=True`）能否 `Add` | **能**，仅内存生效，`Saved = False` |
| 重复 `Add` 同名属性 | **失败** `-2147467259`（必须先 `Delete` 或改 `.Value`） |
| 支持的类型 | 2=Integer、3=Date、4=Text、5=Number、1=Number 全部可用，重开后值不变 |

### 15.4 ⚠️ 对设计的直接影响

**（1）"脏"不能用来判断"用户是否改过数据"。**

既然打标记本身就会弄脏，`wb.Saved` 就**不再是"用户改过"的可靠信号**。

→ **必须改成自己算**：`Base` 快照 vs 当前内容的逐格 diff 才是判定依据
   （这本来就是三方合并的核心，现在多了一个理由）。
→ 永远不要用 `wb.Saved` 决定"要不要提交"。

**（2）打标记的时机必须在 Save 之前、且打完后立即 Save。**

```
Import:  建本地副本 → SaveCopyAs/另存 → Add "MDDT_ROLE"="LOCAL" → wb.Save
```

- `Add` 之后不 `Save` → 下次打开属性就没了（虽然内存里在）
- `Add` 之后用户看到"未保存"，会以为工具做了什么奇怪的事
  → 所以 **Add 完立刻 Save**，让 `Saved` 立刻回到 True

**（3）不要用 `BuiltinDocumentProperties` 做"标记为脏"的手段。**

早期测试脚本曾用 `BuiltinDocumentProperties("Comments")` 把工作簿弄脏来测 Save 耗时。
这确实有效（T18h 确认会弄脏），但也确认了它**改的是真实内容**，会污染文件元数据。
→ 生产代码里**不允许**用这个手段；只在测试脚本里用。

**（4）角色标记可以用，但只作为"辅助识别"。**

- 可靠：能跟着 `SaveAs` / `SaveCopyAs` 走，重开仍在（T18e）
- 局限：**不能反向推断**"有标记 = 有未保存修改"（T18a 反例）
- 局限：重复 `Add` 会失败 → 必须先查存在性或先 `Delete`

推荐读法：

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

**（5）优先用「路径字符串」而不是属性做角色识别。**

`DESIGN.md` §12.5 的 `RoleOf()` 用 `FullName` 比较，**不需要打任何标记**，
因此完全没有"弄脏"副作用。属性标记只用于**跨进程/跨会话**的兜底识别
（例如崩溃后重新打开本地副本时确认"这是我们的副本"）。

## 16. T3 / T5：Save 的写入策略与中断后果（已实测）

> 脚本：`collect/9-t3-t5-driver.ps1`（Windows 驱动）、`collect/9-linux-observer.sh`（Linux 观察者）
> 直方图见：`_smoke/save_strategy.py`（目录变更监听）
> 探针：13.81 MB / 65 sheet / 1.05 M 行的仿真文件（形状与真实文件一致）

### 16.1 ★ 最关键的发现：Excel 使用「临时文件 + 改名」

保存期间对目录做高频轮询，抓到的完整事件序列（相对时间，秒）：

```
 0.000  CREATE   9B370410                    43,843,584   ← 系统临时
 0.771  ==SAVE_START==
 2.008  CREATE   ~$big_probe.xlsx                     0   ← 锁文件
 2.010  MODIFY   ~$big_probe.xlsx                   165
 7.121  CREATE   55790410                             0   ← Excel 自己的临时
 9.609  MODIFY   55790410                    43,843,584
 9.628  MODIFY   55790410                    15,562,381
 9.649  CREATE   ACCB5A11.tmp                14,476,058   ← 原始内容副本
 9.649  DELETE   big_probe.xlsx                          ← ★ 原文件被删除
 9.649  CREATE   big_probe.xlsx              15,562,381   ← ★ 新文件，同一瞬间
 9.649  DELETE   55790410
 9.653  DELETE   ACCB5A11.tmp
11.381  ==SAVE_END==
11.467  DELETE   ~$big_probe.xlsx
```

**结论：**

| 事实 | 含义 |
|---|---|
| 新内容先写进临时文件 `55790410` | 原文件在写入期间**完全未被触碰** |
| 写完后 `DELETE` 旧文件 + `CREATE` 新文件在**同一毫秒级** | 这是 **rename 语义**，不是就地覆写 |
| 中间没有"半个文件"的状态 | ✅ **对读者是原子的** |
| `ACCB5A11.tmp` 保存了原始内容（14,476,058 = 原大小） | 这是 Excel 的回滚副本 |

### 16.2 杀进程实验：Save 窗口内杀掉，文件完好

`Save()` 在 13.81 MB 上耗时 **4.2–5.8 s**（与真实文件 6.77 s 同量级）。

| 阶段 | Save 耗时 | 杀点是否落在窗口内 | zip 完好 | Excel 能打开 | A1 内容 | 判定 |
|---|---|---|---|---|---|---|
| P1 正常保存 | 4178 ms | — | ✅ 74 项 | ✅ | `P1` | ok |
| **P3 杀 @500 ms** | 2854 ms（抛 `0x800706BE`） | **是** | ✅ | ✅ | `REFERENCE_DIRTY`（**旧内容**） | **ok** |
| **P4 杀 @1500 ms** | 3681 ms（抛 `0x800706BE`） | **是** | ✅ | ✅ | `P4_KILL_1500MS dirty`（新内容） | **ok** |
| P5 杀 @3000 ms | 4444 ms | 否（太晚） | ✅ | ✅ | 新内容 | ok |
| P6 杀 @5000 ms | 5800 ms | 否（太晚） | ✅ | ✅ | 新内容 | ok |
| P7 保存后杀（对照） | — | — | ✅ | ✅ | `P7` | ok |

**关键的两行：**

- **P3（500 ms 杀）：文件回落到旧内容且完全可读。** 说明写入期间原文件真的没被破坏。
- **P4（1500 ms 杀）：文件是完整的新内容。** 说明 rename 已完成但 Excel 还没返回。

→ **`Save()` 中断不会损坏服务器文件。** 最坏情况是"改动没生效"，不是"文件坏了"。

### 16.3 残留文件

每轮结束后 `stray=[]`，无 `~$` 残留、无 `.tmp` 残留。
杀进程也**没有**留下需要清理的临时文件（被杀时临时文件由操作系统回收或被后续 rename 覆盖）。

### 16.4 无 `~$` 残留是否会阻塞下一次保存

未观察到阻塞。崩溃后重新打开 / 保存均正常。

**但注意**：本测试杀的是"持有 `~$` 的那个进程"。
`~$` 文件的删除发生在正常 Close 路径，杀进程时它可能残留 → 见 §16.6 待补。

### 16.5 对设计的直接影响

| 原设计假设 | 实测结论 | 动作 |
|---|---|---|
| "Save 可能损坏文件 → 需要提交前备份" | **不需要**。rename 语义保证原文件在写入期间不被触碰 | ✅ **维持"不做提交前备份"**，省下 106–245 s |
| "Save 有 ~7 s 风险窗口" | 窗口存在，但**风险是"改动未生效"，不是"文件损坏"** | 失败只需重试 |
| "外部读者（SAS）可能读到半个文件" | Windows 侧观察**不成立**（见 §16.6 需 Linux 侧确认） | 待 T5 |

**这是本轮最有价值的结果：它确认了一个我们原本担心的设计决定是正确的。**
"不做提交前备份"不需要推翻 —— 因为 Excel 本来就在用原子 rename。

但必须补一条**新规则**：

> **失败恢复 = 重试，不是修复。**
> 若 `Save()` 抛错（如 `0x800706BE` RPC 失败），状态未知。
> 正确做法：**重新打开服务器文件、核对内容、再决定**。
> 绝不能假设"保存成功了"而释放锁。

### 16.6 尚未覆盖的三项（需在工作环境跑）

| # | 待测 | 为什么需要 |
|---|---|---|
| **T5** | Linux 侧在 rename 瞬间读到什么 | Windows 侧轮询有 ~1 ms 分辨率，可能漏掉"文件消失"的瞬间。Linux 侧 `stat` 会返回 `ENOENT` → **SAS 会报错** |
| **T3b** | 杀进程后 `~$` 是否残留、是否阻塞 | 本测试观察窗口太短 |
| **T3c** | 网络中断（拔线/RDP 断）而非杀进程 | 更接近真实场景，且可能产生半写状态 |

**其中 T5 最重要**：rename 的 1 ms 窗口对 Windows 观察者不可见，但 Linux 的 `open()` 可能正好落在这个窗口里。

### 16.7 ★ T3-on-share：在真实 SMB 共享上复测（已完成）

本地盘上的结论必须在共享上复验，否则不能作为架构依据 ——
Excel 对网络位置**可能**走不同的保存策略。以下为真实共享上的结果。

**探针**：真实 `book.xlsx`，11,468,288 字节（10.94 MB），16 个 zip 条目。

| 阶段 | 文件大小 | zip 完好 | Excel 能开 | 杀点落在窗口内 | 判定 |
|---|---|---|---|---|---|
| P1 正常保存 | 11,468,288 | ✅ | ✅ | — | ok |
| **P3 杀 @500 ms** | 11,468,313 | ✅ | ✅ | **是** | **ok** |
| **P4 杀 @1500 ms** | 11,468,313 | ✅ | ✅ | **是** | **ok** |
| **P5 杀 @3000 ms** | 11,468,313 | ✅ | ✅ | **是** | **ok** |
| **P6 杀 @5000 ms** | 11,468,313 | ✅ | ✅ | **是** | **ok** |
| P7 保存后杀（对照） | 11,468,292 | ✅ | ✅ | 否 | ok |

`>>> ALL PHASES OK: the server file was never left damaged.`

**与本地盘一致的三点：**

1. **rename 语义成立** —— 四轮杀进程（含真正落在 Save 窗口内的 P3–P6）之后，
   每个阶段 `zipOk=True` 且 `excelOpens=True`，没有任何一轮留下损坏文件。
2. **文件没有被截断** —— 所有大小都在 11,468,2xx–11,468,313 之间，波动仅 25 字节
   （zip 内部时间戳差异），没有出现"半写"尺寸。
3. **Save 耗时量级一致** —— 11 MB 文件在共享上同样是数秒级窗口，杀点可稳定命中。

→ **"不做提交前备份"这个决定在 SMB 共享上同样成立。**
→ 本地盘结论可直接推广到生产路径。

### 16.8 ⚠️ 新现象：共享上崩溃会残留 .tmp 文件

这是本地盘上**没有观察到**的差异：

```
P7_KILL_AFTER_OK ... stray=[52C6A2C6.tmp, EE6B8C50.tmp]
```

本地盘首轮测试时 `stray=[]`，共享上出现两个残留：

| 现象 | 说明 |
|---|---|
| 残留文件名是 **8 位十六进制 + .tmp** | 与 §16.1 观测到的 `ACCB5A11.tmp` **同一命名规则** |
| 只出现在 **P7（保存完成后杀）** | 保存中途杀（P3–P6）反而没残留 |
| 其余阶段 `stray` 为空 | 说明正常路径会自行清理 |

**为什么 P7 才残留：** P7 在 `Save()` **返回之后**立即杀进程。
此时 Excel 已发布新文件，但内部清理（删除回滚副本）尚未执行完。
本地盘首轮没抓到，可能是时序竞争 —— 不是本地盘不会残留，而是**观察窗口太短没撞上**。

**影响评估：**

| 问题 | 结论 |
|---|---|
| 会损坏文件吗 | **不会**，此阶段 `zipOk=True / excelOpens=True` |
| 会阻塞下一次保存吗 | **不会**，Excel 用随机名字，不会撞名 |
| 会累积吗 | **会**。每次崩溃留 1–2 个，长期堆在共享目录里 |
| 会被工具误认成数据文件吗 | **不会**，扩展名不是 `.xlsx` |
| 会被 SAS 脚本扫到吗 | **风险点**，见下 |

**必须加一条清理规则（新增 T20）：**

> 工具启动时（或每次提交前）扫描共享目录，删除 **超过 24 小时**的
> `????????.tmp` 文件（8 位十六进制 + `.tmp`）。
> **不要删当天的** —— 可能是别的用户正在保存。

```vba
' 只删"老"的 8 位十六进制 .tmp，绝不删今天的
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

**对 SAS 的提示**：如果 SAS 脚本用通配符扫目录（`*.xlsx` 不受影响，
但 `for f in dir/*` 或 `ls | grep` 会命中），需明确排除 `.tmp`。
→ 这属于**告知项**，不是工具缺陷。

### 16.9 T3 结论汇总：可以对 SAS 说什么

| 读者遇到的时刻 | 结果 | SAS 需要改吗 |
|---|---|---|
| rename 之前 `open()` | 完整**旧**内容 | 否 |
| rename 之后 `open()` | 完整**新**内容 | 否 |
| 正好卡在那一毫秒 | `ENOENT` 报错 | 需要**重试一次** |
| `actimeo=1` 缓存未过期 | 最多 1 秒看到旧版本 | 否（与今天一致） |
| **撕裂读（读到半个文件）** | **不存在** | — |

**"撕裂"在 rename 语义下不会发生** —— 读者永远拿到某一个完整 inode。
最坏情况是一次 `ENOENT`，重跑 SAS 即可。用户已确认**可接受**。



## 17. Excel API 陷阱与测试驱动陷阱（本轮踩到，全部已定位）

> 这些不是产品缺陷，但**每一条都会让人误判**，其中前三条会让测试得出错误结论。

### 17.1 ❌ `Range.Find(What:="*")` 会静默返回错误结果

测 `SafeLastCol` / `SafeLastRow` 时，对 7 张表（最宽 9 列、最长 11 行）调用：

```vba
ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
```

**结果：7 张表全部返回 `col = 2`**，无论真实宽度是 2、5、7、9。
改用 `ws.UsedRange.Find(...)` 也一样返回 `$B$1`。

| 方法 | 结果 |
|---|---|
| `ws.Cells.Find` | ❌ 全部 col=2 |
| `ws.UsedRange.Find` | ❌ 全部 $B$1 |
| `UsedRange.Columns.Count` | ✅ 正确（7 / 5 / 2 / 3 / 6 / 9 / 2） |
| 读 `Row 1` 数组扫描 | ✅ 正确 |
| 读第 1 列数组向上扫描 | ✅ 正确 |

→ **禁用 `Find` 求边界。** 它不报错，直接给错答案。

### 17.2 ❌ `UsedRange` 在真实文件上不可用（幽灵行）

已知事实（`FINDINGS.md` 测试 2）：一张 sheet 被意外撑到 **1.1 M 行 × 2 M 列**。
`UsedRange` 会把整个虚假维度当数据。

→ **边界必须用「读数组 + 扫描非空」**，可容忍幽灵维度：

```vba
' 列边界：读第 1 行 4096 格，1 次 COM 调用
arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 4096)).Value2
For i = 1 To 4096
    If Not IsCellEmpty(arr(1, i)) Then n = i
Next i
```

```vba
' 行边界：窗口倍增。命中窗口边界就翻倍，未命中即真实末行
win = 8192
Do
    arr = ws.Range(ws.Cells(1, col), ws.Cells(win, col)).Value2
    For r = win To 1 Step -1
        If Not IsCellEmpty(arr(r, 1)) Then n = r: Exit For
    Next r
    If n <> win Then Exit Do          ' 真正的末行
    If win >= 400000 Then Exit Do     ' 上限保护
    win = win * 2
Loop
```

> 幽灵维度只让扫描变慢（几次 COM 调用），**不会截断真实数据**。

### 17.3 ❌ `Dim` 在循环里不会重新初始化

VBA 把 `Dim` 提到**过程作用域**，循环内的 `Dim` **只初始化一次**。

```vba
For i = 1 To 7
    Dim dataRows As Long          ' 只在 i=1 时是 0
    For r = 1 To 100
        dataRows = dataRows + 1   ' i=2 时从上一轮的 100 继续累加
    Next r
    Debug.Print dataRows          ' 100, 200, 300 ... 全错
Next i
```

**实测表现**：第一张表 10 行，第二张表报 15 行（10+5），第三张 15，第四张 20。
所有百分比、去重计数**全部污染**。

→ 循环体开头**必须显式重置**每个累加量：

```vba
ghostNote = "": firstAMG = 0: nProduct = 0: nAnchor = 0
nAMG = 0: nCollide = 0: dupRows = 0: blankRows = 0: dataRows = 0
dupPct = 0
```

### 17.4 ❌ 参数名会遮蔽同名数组的索引表达式

```vba
Private Function IsEmptyV(ByVal v As Variant) As Boolean   ' 参数叫 v
    ...
End Function

arr = ws.Range(...).Value2
For i = 1 To 4096
    If Not IsEmptyV(v(1, i)) Then n = i     ' ← 灾难
Next i
```

修复前 `n` 永远是 0，表现为"每张表都没有表头"。
把参数改名为 `cell`、数组改名为 `arr` 后正常。

→ 规则：**参数名与调用方的数组变量名不得相同。**

### 17.5 ⚠️ 测试驱动陷阱：`Run` 会在错误的工作簿里找宏

```python
app.Run("ZZRUNSCHEMA")          # ❌
```

若 VBA 过程内部会 `Workbooks.Open(target)`，打开后**活动工作簿变成 target**，
Excel 便到 target 里找 `ZZRUNSCHEMA` → 报

```
Cannot run the macro 'ZZRUNSCHEMA'.
The macro may not be available in this workbook or all macros may be disabled.
```

这个报错**极易被误读成"VBA 编译失败"**。

正确写法（二选一）：

```python
app.Run("'" + xlsm_path + "'!ZZRUNSCHEMA", target)   # 限定工作簿
```

或让过程接收路径参数、**全程不切换活动工作簿**。

> 附带确认：`VBComponents.Add(1)` + `AddFromString` 在连续多次后运行会失败；
> `VBComponents.Import(bas_file)` 稳定（与已验证的 `run_diffperf.py` 一致）。
> → **驱动统一用 `Import`。**

### 17.6 ⚠️ 编译探针的判定意义

调用**任何** Public 过程都会强制编译整个项目。所以：

| 现象 | 结论 |
|---|---|
| `ZZCOMPILEPROBE` 打印 `compile ok` | **整个模块编译通过**，不存在"变量未定义" |
| 随后 `Run` 报"变量未定义" | 错误**不在这段代码**，而是驱动找错了工作簿（§17.5） |

本轮我在 `HeaderLastCol` 上反复找"未定义变量"，实际根因是 §17.5。
**编译探针通过后，不应再去怀疑语法。**

## 18. 对设计的修订（基于以上实测）

1. **写回用 `Range.Copy`，明确不用 `Value2`** —— 已在 `PROCESS.md §4` 写明，此处获得实测支撑
2. **新增：Copy 后显式复制行高/列宽**（3 行代码）
3. **新增：清理范围 = Union(服务器区域, 本地区域)** —— 否则删行失效
4. **修订：不用 VBA `Name` 做重命名换页** —— 用 `MoveFileExW` 或 FSO `MoveFile(OverWriteFiles:=True)`
5. **新增：key 列文本格式必须在赋值前设置**
6. **新增：哈希用上述 `Double` 分段 FNV-1a**（不可用 `Mod`、不可用 `Long` 直接乘）
7. **新增：错误码判定要同时接受 5 和 32**
