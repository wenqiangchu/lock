# 验证与数据采集清单

> 目标：把 DESIGN.md 里所有"假设"变成"实测数据"。
> 原则：**任何脚本都不在生产文件旁边写东西。** 所有共享盘写入都收敛到你显式指定的探针目录。

---

## 0. 编码与语言约定

- **`.ps1` / `.bas` 全部纯 ASCII 英文**（无中文、无 emoji、无框线符号）
  原因：PowerShell 5.1 在文件无 BOM 时按系统 ANSI 代码页（中文 Windows 上是 GBK）读取，
  任何非 ASCII 字节都会变成乱码，甚至破坏解析。
- **`.md` 文档用中文**，这些文件不会被程序解析，编码随意。
- 如果你要改脚本，请保持 ASCII。不要往 `.ps1` / `.bas` 里加中文注释或 emoji。
- 回答问卷、写备注可以用中文。

---

## 1. 先看：每个脚本会碰什么（重要）

| 脚本 | 会写共享盘吗 | 具体位置 |
|---|---|---|
| `1-office-env.ps1` | **不会** | 桌面报告 + `%LOCALAPPDATA%` 里的 profile 标记文件 |
| `2-file-profile.ps1` | **不会** | 复制到 `%TEMP%` 解压，用完删除 |
| `3-measure.ps1` | **仅当传了 `-WorkDir`** | `<WorkDir>\_exs_probe_<用户>\`（自建，可选自动删除） |
| `4-test-lock.ps1` | 会，在你给的 `-Path` | 一个锁文件 + 一个探针目录 |
| `5-test-replace.ps1` | 会，但只在 `-Dir` 下自建的子目录里 | `<Dir>\_exs_probe_<用户>_swap\` |
| `4-vba-lock.bas` | **不会** | `%TEMP%` 和桌面 |
| `6-test-copyfidelity.bas` | **不会** | `%TEMP%` 和桌面 |

**`1`、`2`、`6` 可以直接对着生产文件跑，安全。**

---

## 2. 写测试该用哪个目录（关于 home folder）

**不要用 home folder 替代项目共享盘来测"锁"和"重命名"。**

原因：这两个测试测的是**某台 Samba 服务器上某个 share 的配置**（`strict rename`、
`delete readonly`、oplocks、ACL 继承）。换个 share 结论可能完全不同。

推荐布局：

```
\\SERVER\SHARE\book.xlsx                  <- 生产文件，这些脚本从不写它
\\SERVER\SHARE\_exs_probe_<你>\           <- 你的探针目录，测完删掉
```

如果项目共享盘里不允许你建目录：
- 找 IT 在同一 share 上开一个临时目录
- 实在不行，**只有 `3-measure.ps1`（纯计时）** 可以退到 home folder 跑，
  但要在报告里明确标注"锁/重命名结论不适用于项目 share"

`4-test-lock.ps1` 要求**两台机器都能写同一个路径**。如果你的 home 是私有的，
就必须改用共享的探针目录。

---

## 3. 准备

1. 把整个 `collect` 文件夹拷到工作电脑，例如 `C:\Users\<你>\Desktop\collect\`
2. 打开 **PowerShell**（不需要管理员权限）
3. 如果脚本被执行策略拦住：
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
4. 每个脚本都会在**桌面**生成一个 `exs_*.txt` 报告，把这些回传给我
5. 把下面的 `\\SERVER\SHARE` 占位符换成你的真实路径

---

## 4. 执行顺序

| # | 脚本 | 耗时 | 需要 Excel | 需要同事 | 阻塞什么 |
|---|---|---|---|---|---|
| 1 | `1-office-env.ps1` | 1 分钟 | 否 | 否 | 宏策略 / 受信任位置 → 能不能部署 |
| 2 | `2-file-profile.ps1` | 2 分钟 | 否 | 否 | 文件结构 → 能否整块 Copy、key 是否可用 |
| 3 | `3-measure.ps1` | 5 分钟 | **是** | 否 | 提交窗口时长、锁开销 |
| 4 | `4-test-lock.ps1` | 3 分钟 | 否 | **是** | 互斥锁方案成立性 |
| 5 | `5-test-replace.ps1` | 5 分钟 | **是** | 可选 | 能否升级成"重命名换页"架构 |
| 6 | `6-test-copyfidelity.bas` | — | — | — | ✅ **已在本机验证，跳过**（见 `../VERIFIED.md`） |

1、2 立刻做（零风险）。3、5 挑非高峰期做。4 找同事腾 3 分钟。

**第 6 项已完成**：超链接、填充色、字体色、合并格、批注、数据验证、条件格式
都能被 `Range.Copy` 保留；`Value2` 会全部丢失；`Range.Copy` 不携带行高/列宽。

---

## 5. 第 1 步：环境与策略

```powershell
cd "$env:USERPROFILE\Desktop\collect"
.\1-office-env.ps1
```

看什么：

| 字段 | 关心 | 影响 |
|---|---|---|
| `ClickToRun Platform` / `EXCEL.EXE version` | x86 还是 x64 | `Declare PtrSafe` 与 `LongPtr` 写法 |
| `VBAWarnings` | 1 / 2 / 3 / 4 | **4 = 只允许签名宏** → 必须先解决证书 |
| 受信任位置列表 | 有没有你的共享盘 | 没有则需 IT 添加 |
| `AllowNetworkLocations` | 是否 1 | 为 0 则 UNC 路径无法加为受信任位置 |
| 杀毒 / 实时扫描 | 有无 | 影响 Save 偶发失败 → 决定重试策略 |
| `SESSIONNAME` / `LOCALAPPDATA` | 是否 RDP | 本地副本落盘位置 |

**profile 持久性验证**（重要）：

1. 跑一次 `.\1-office-env.ps1`（它会写一个标记文件）
2. **注销，重新登录**
3. 跑 `.\1-office-env.ps1 -CheckProfile`

标记文件不见了 = `%LOCALAPPDATA%` 不持久，本地工作副本会丢，设计需要改。

---

## 6. 第 2 步：文件结构画像（只读）

```powershell
.\2-file-profile.ps1 -File "\\SERVER\SHARE\book.xlsx"
```

复制到 `%TEMP%`、解压、解析 XML。不打开 Excel，所以不产生 `~$`、不占用共享盘。

这一份报告基本能回答所有悬而未决的问题：

| 字段 | 决定什么 |
|---|---|
| `dimension` vs 实际 `cells` | 10MB 是不是格式垃圾 → 能否瘦身 |
| 每 sheet 的 `rows` / `cells` | 真实规模 → 提交窗口估算 |
| `hyperlink` 数量与 `location` | TOC 用单元格超链接还是 `HYPERLINK()` 公式 |
| `tableParts`（ListObject） | 有则不能整块 Clear，会破坏表 |
| `mergeCell` 数量 | 有合并格 → key 列可能为空 |
| `dataValidation` / `conditionalFormatting` | 决定 Clear 的激进程度 |
| `sheetProtection` | **有保护则写回直接失败** |
| `fileSharing` | 旧版"共享工作簿"与整个设计冲突 |
| A 列统计 | key 是否真的唯一 / 是文本 / 有无前导零 |
| `cellXfs` 数量、`styles.xml` 大小 | 样式碎片化 → 瘦身收益 |
| `xl/media`、`xl/charts`、`pivotCache` | 有没有我没算到的东西 |
| `externalLinks`、`connections` | 外部依赖 |
| `calcPr` | 手动计算会导致值不刷新 |

**报告里出现 `WARN` 就先停下**，把报告发我再继续。

---

## 7. 第 3 步：网络与 Excel 耗时

**这个脚本默认不写共享盘。** 只有传 `-WorkDir` 才会在探针目录里写：

```powershell
# 纯网络计时，完全不写共享盘
.\3-measure.ps1 -File "\\SERVER\SHARE\book.xlsx"

# 完整测试（含 Excel 打开/保存），只在探针目录里写
.\3-measure.ps1 -File "\\SERVER\SHARE\book.xlsx" -WorkDir "\\SERVER\SHARE\_exs_probe_你的名字"

# 完全跳过 Excel
.\3-measure.ps1 -File "\\SERVER\SHARE\book.xlsx" -SkipExcel
```

指标含义：

| 指标 | 预期 | 决定 |
|---|---|---|
| `Copy-Item` 整个文件 | 0.3–2 s | 网络本身是否瓶颈 |
| `stat` ×20（存在+长度+时间） | 每次 1–10 ms | 锁记账的开销 |
| 小文件 建+写+删 | 每次 5–50 ms | `owners.tsv` 轮询频率上限 |
| Excel 打开本地副本 | 5–10 s | **这是 Excel 解析时间，不是网络** |
| Excel 打开共享盘副本 | 上者 + 网络 | 差值 = 真实网络成本 |
| Excel Save 到共享盘 | 3–10 s | **= 提交时占用窗口的主体** |
| `~$` 文件出现/消失 | 是/是 | 孤儿锁风险 |
| 只读打开是否产生 `~$` | 期望否 | 「只读查看」按钮能否真正不占用 |

判定：`Excel Save 到共享盘` ≤ 10 s → 方案成立（现状是小时级）。> 30 s → 先瘦身。

---

## 8. 第 4 步：互斥锁验证（需要同事，3 分钟）

**两台机器必须用同一个、都能写的路径。**

A 侧（你）：
```powershell
.\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_你的名字\lock.tst" -Mode Hold -Seconds 90
```
窗口保持 90 秒不关。

B 侧（同事）：
```powershell
.\4-test-lock.ps1 -Path "\\SERVER\SHARE\_exs_probe_你的名字\lock.tst" -Mode Try
```

期望：B 侧报告 `SHARING VIOLATION → 锁生效`。

| 结果 | 含义 | 影响 |
|---|---|---|
| B 被拒绝 | SMB 服务端强制互斥 | 锁方案成立 |
| B 也成功 | 该共享未强制 | 改用 mkdir 原子锁（+0.5 天） |

脚本会**拒绝**往"看起来像生产目录"的地方写（含 `.xlsx` 但路径里没有 `_exs_probe`），
除非你加 `-Force`。

可选交叉验证：A 侧跑 `4-vba-lock.bas` 里的 `HoldLockVBA`，B 侧再跑 `-Mode Try`，
证明 VBA 的 `Open Lock` 也参与同一套机制。

---

## 9. 第 5 步：重命名换页可行性（架构升级的钥匙）

如果通过，**提交时完全不锁文件**：读者可以一直开着，提交窗口从 20 s 降到约 2 s。

```powershell
# 1) 准备（脚本会自己在 -Dir 下建 _exs_probe_<你>_swap 子目录）
.\5-test-replace.ps1 -Mode Prepare -Dir "\\SERVER\SHARE\_exs_probe_你的名字"

# 2) 从探针目录用 Excel 打开 a.xlsx，【保持打开】，读写方式

# 3) 尝试替换
.\5-test-replace.ps1 -Mode Replace -Dir "\\SERVER\SHARE\_exs_probe_你的名字"

# 4) 关闭 Excel 后
.\5-test-replace.ps1 -Mode Verify  -Dir "\\SERVER\SHARE\_exs_probe_你的名字"
.\5-test-replace.ps1 -Mode Cleanup -Dir "\\SERVER\SHARE\_exs_probe_你的名字"
```

怎么读结果：

| 结果 | 含义 | 决策 |
|---|---|---|
| 文件开着时 `SUCCESS(0)` | Samba 允许覆盖被打开文件 | **采用重命名换页架构** |
| `SHARING_VIOLATION(32)` | 不允许 | 走保守方案 `Workbook.Open` + `Save` |
| `ACCESS_DENIED(5)` | 可能是"被占用"，也可能是没写权限 | 关掉 Excel 重跑；重跑成功 = 确实是被占用 |
| 替换成功、Excel 无报错、内容仍是 OLD | 符合 POSIX rename 语义 | 正常 |
| 替换成功但 Excel 报文件损坏 | 危险 | **不要用这个架构** |

注意：本地 NTFS 上 Windows 返回的是 **5** 而不是 32。判定要同时接受 5 和 32，
并用"释放后重试"这一步来确认。

---

## 10. 第 6 步：跨工作簿 Copy 保真性（已完成，跳过）

已在开发机用 Excel 16.0 x64 跑通，完整结果见 `../VERIFIED.md`。

可选：对着**真实文件的副本**再跑一次，确认实际 TOC 超链接形态也没问题。
做法：把 `TestCopyFidelity` 里的源工作簿换成你的真实文件，保持
`SHOW_MSGBOX = True`，然后看报告 C 段。

---

## 11. 已知的 Excel/VBA 坑（已踩过并解决）

一开始就用正确写法，细节见 `../VERIFIED.md`。

| 坑 | 症状 | 正确写法 |
|---|---|---|
| 用 VBA `Mod` 做哈希乘法 | `Run-time error 6 Overflow` | 用 `Double` 分段做 32 位模乘 |
| `Name dst` 覆盖已存在文件 | `error 58 File already exists` | `MoveFileExW(REPLACE_EXISTING)` 或 FSO `MoveFile(OverWriteFiles:=True)` |
| 先赋值再设 key 列文本格式 | `"00123"` 变 `123` | 先设 `NumberFormat = "@"`，再赋值 |
| 覆盖较小区域前不清除 | 旧行残留，看起来像删不掉 | Clear `Union(服务器区域, 本地区域)` |
| 隐藏窗口里 `Range("A1").Select` | `error 1004` | 不要 Select；必要时先 `Activate` 且 `Visible=True`，并加错误保护 |
| `Kill` / `MkDir` 无保护 | `error 53` / `error 75` | 先用 `Dir$()` 判断，或 `On Error Resume Next` 包住 |
| `MoveFileEx` 错误码不一致 | 本机 5，Samba 可能是 32 | 同时接受 5 和 32 |
| 公开过程名以 `__` 开头 | 能编译但 `Application.Run` 找不到 | 不要用双下划线开头的公开过程名 |
| `Dim x As Variant` 后 `ReDim x(...)` | 编译错误 | 写 `Dim x()` 或去掉 `Dim` |
| 重复 `Option Explicit` | 编译错误 | 每个模块只留一处 |
| VB 字符串里嵌 ASCII 双引号 | 语法错误 | 转义成 `""`，或改用【】等其它符号 |

---

## 12. 回传清单

- [ ] `Desktop\exs_office.txt`
- [ ] `Desktop\exs_fileprofile.txt`
- [ ] `Desktop\exs_measure.txt`
- [ ] `Desktop\exs_lock_<用户名>.txt`（两台机器各一份）
- [ ] `Desktop\exs_replace.txt`（Prepare 与 Replace 两次）
- [ ] `7-questions.md` 的答案
- [ ] `-CheckProfile` 复跑结果

---

## 13. 清理

```powershell
Remove-Item "\\SERVER\SHARE\_exs_probe_你的名字" -Recurse -Force
```

`3-measure.ps1` 删除探针目录前会问你。
`5-test-replace.ps1 -Mode Cleanup` 会清掉自己建的文件。

---

## 14. 决策矩阵

| 发现 | 结论 |
|---|---|
| `VBAWarnings = 4` 且无签名能力 | 部署受阻，先解决证书 |
| `AllowNetworkLocations = 0` | 走本地中转，或申请策略例外 |
| 任一 sheet 有 `sheetProtection` | 需要解锁流程，或改设计 |
| 有 `fileSharing`（旧共享工作簿） | 必须先转普通工作簿 |
| 有 `tableParts`（ListObject） | 不能 Clean 表，只覆盖数据区 |
| 重命名换页成功 | 采用，提交窗口约 2 s，读者永不阻塞你 |
| `Excel Save` > 30 s | 先瘦身文件再继续 |
| profile 非持久 | 本地副本方案必须改 |
| 锁不生效 | 改用 mkdir 原子锁 |