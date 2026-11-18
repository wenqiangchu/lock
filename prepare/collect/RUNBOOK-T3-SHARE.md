# T3-on-share 操作清单

> 目的：确认「Excel 用临时文件 + rename 保存」这个结论，在 **SMB 共享**上也成立。
> 我之前的 T3 是在本地 NTFS 盘上跑的，严格说没验过共享。
> 只需要 Windows + 一个 PowerShell 窗口，**不需要 PuTTY，不需要 Linux**。
> 耗时约 **10 分钟**，其中 5 分钟是等待。

---

## 现在做的事（4 步）

### 第 1 步：准备探针目录（一次性，1 分钟）

```powershell
# 目录名必须含 _probe —— 脚本会硬校验，不符合直接拒绝
New-Item -ItemType Directory -Force '\\你的服务器\共享\_mddt_probe'

# 确认探针目录已建好
Test-Path '\\你的服务器\共享\_mddt_probe'
```

**不需要手动复制文件** —— 脚本会自己从 `-SourceFile` 复制。

---

### 第 2 步：⚠️ 关闭你正在编辑的 Excel（重要）

脚本要**杀掉**它自己创建的 Excel 进程。它靠"启动前/启动后的 PID 差集"识别自己那个实例。

**如果你同时开着别的 Excel 窗口，有误杀未保存工作的风险。**

```powershell
# 确认没有 Excel 在跑
Get-Process EXCEL -ErrorAction SilentlyContinue
```

有输出就先保存关闭。没输出就可以继续。

---

### 第 3 步：运行（10 分钟，含 5 分钟自动等待）

```powershell
cd G:\projects\mddt_lock_issue

.\collect\9-t3-t5-driver.ps1 `
    -ProbeDir   '\\你的服务器\共享\_mddt_probe' `
    -SourceFile '\\你的服务器\共享\book.xlsx' `
    -SkipReaderHold
```

**参数说明：**

| 参数 | 为什么这样填 |
|---|---|
| `-ProbeDir` | 探针目录，**名字必须含 `_probe`** |
| `-SourceFile` | 用**真实的 book.xlsx**（11.44 MB）。⚠️ 别用小文件 |
| `-SkipReaderHold` | 跳过 P2（那个阶段需要 Linux 侧配合）。**跳过它不会等 20 秒倒计时** |

**跑起来后你什么都不用做**，等它自己跑完。中间会看到：

```
PHASE >>> P1_NORMAL_SAVE        (正常保存，约 5-8 秒)
PHASE >>> P3_KILL_500MS         (500ms 时杀 Excel)
PHASE >>> P4_KILL_1500MS        (1500ms 时杀)
PHASE >>> P5_KILL_3000MS
PHASE >>> P6_KILL_5000MS
PHASE >>> P7_KILL_AFTER_OK      (对照组)
```

---

### 第 4 步：把最后 30 行贴回来

脚本跑完会打印一张 SUMMARY 表。**把终端最后 30 行复制给我。**

也可以从日志文件里取（脚本结束时会告诉你路径）：

```powershell
Get-Content 'C:\Users\<你>\AppData\Local\Temp\mddt_t3t5_*.log' |
    Select-Object -Last 30
```

---

## 我要看什么（4 个数）

贴回来后我只检查这 4 项：

| # | 看哪里 | 期望值 | 不对意味着什么 |
|---|---|---|---|
| 1 | `source size` 那一行 | **≈11.44 MB**（约 11,000,000+ 字节） | 如果只有几 MB → 探针太小，Save 太快，杀不中窗口，**测试无效** |
| 2 | 各阶段 `Save() returned after NNNN ms` | **3000–8000 ms** | 只有几十 ms → 同上，测试无效 |
| 3 | SUMMARY 表每行 `zipOk` / `excelOpens` | 全部 `True` / `True` | 出现 `False` → **文件被写坏了，必须加备份** |
| 4 | SUMMARY 表 `killInWindow` 列 | **P3、P4 至少有一个是 `True`** | 全是 `False` → 一次都没杀进窗口，**等于没测** |

第 4 项特别重要：`killInWindow=False` 说明杀进程发生在 Save 已经返回之后，那这一轮**什么都证明不了**。

---

## 顺带想知道的（可选，2 秒）

脚本跑完后，共享上那个 `_mddt_probe` 目录里还有没有残留文件？

```powershell
Get-ChildItem '\\你的服务器\共享\_mddt_probe' -Force | Select-Object Name, Length
```

正常情况下应该**只有 `_phase.txt`**（探针文件脚本会自动删掉）。

如果出现 `~$xxx.xlsx` 或 `*.tmp` → 说明崩溃后会残留垃圾，我需要加清理逻辑。

---

## 安全边界（为什么不会碰你的生产文件）

| 机制 | 说明 |
|---|---|
| 探针目录名校验 | 不含 `_probe` → **硬拒绝**（我实测触发过） |
| 同目录校验 | 探针目录 = 源文件目录 → **硬拒绝** |
| 源文件只读 | 脚本只 `Copy-Item`，从不以可写方式打开源文件 |
| 每阶段重建 | 每轮从本地 pristine 副本重新复制探针 |
| 只杀自己的进程 | 按 PID 差集识别，不碰你手动开的 Excel |
| 结束清理 | 默认删除探针文件（`-KeepProbe` 才保留） |

**唯一需要你配合的**：第 2 步先关掉自己的 Excel。

---

## 可能遇到的提示

| 提示 | 处理 |
|---|---|
| `REFUSED: ProbeDir must contain '_probe'` | 目录名改成含 `_probe` |
| `REFUSED: ProbeDir is the same directory as SourceFile` | 探针目录不要放在 book.xlsx 同一个目录 |
| `SourceFile does not exist` | 路径写错了，注意 UNC 要用 `\\服务器\共享\` 开头 |
| 卡住不动超过 2 分钟 | 可能在某轮 Save 上挂了，`Ctrl+C` 后把已输出的部分贴我 |
| `could not identify EXCEL pid` | 有别的 Excel 在跑干扰了，关掉重跑 |

---

## 我为什么不直接在共享上跑

我的开发机连不到你们的共享。而且这个测试**必须在你的真实网络环境里跑**才有意义 ——
之前的废弃数据就是因为从家里经 WAN 测的（延迟 3090 ms vs 现在的 2 ms，差 1500 倍）。

---

## 跑完之后

| 结果 | 下一步 |
|---|---|
| 4 项都符合 | ✅ **测试阶段结束，可以开工写代码** |
| 出现 `zipOk=False` | 🔴 设计要加"提交前备份"，代价 106–245 秒/次，需要重新权衡 |
| `killInWindow` 全 False | 🟡 探针太小或机器太快，需要用更大的文件重跑 |

---

## 跑的过程中我同时在做

- **T19**：锁的 fencing 回读校验（`MkDir` → 写 owner.txt → 回读确认 token 是自己的）
- **`MdlMerge` 骨架 + 单元测试**：纯算法，不依赖 Excel，不需要工作机