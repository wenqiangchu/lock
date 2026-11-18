# T5 / T3 现场操作手册

> 你需要做的只有 **第 1、2、3 步**。第 4 步是我做的。
> 全部只读，唯一写入是探针副本，结束后自动删除。

---

## 现在就要做的事（按顺序）

### 第 1 步：在 PuTTY 里跑预检（2 分钟，不需要等 Windows）

目的：一次性拿到所有"路径不对就会白跑"的信息。

```bash
cd /path/to/collect          # 把这两个脚本传上去
chmod +x 9-preflight.sh 9-linux-observer.sh

./9-preflight.sh /path/to/_mddt_probe/big_probe.xlsx
```

**如果还不知道路径**，先跑这两条：

```bash
mount | grep -i -E 'cifs|smb'
ls /mnt
```

**把整页输出贴回来给我。**

我主要看 5 个数：

| 看什么 | 为什么 |
|---|---|
| 挂载路径 + `cache=` 选项 | `cache=loose` 会让你读到过期字节，测试就白做了 |
| `unzip -t` 是否 PASS | 确认 Linux 侧能正确解析这个 xlsx |
| 大小写敏感性 | 决定锁名必须全大写是否必要 |
| `second mkdir` 是否被拒绝 | 锁原语在 Samba 上是否工作 |
| 单次读延迟 | 判断当前是否在 WAN 上（之前你从家测出 3090 ms） |

> ⚠️ 第 7 项如果在**几十毫秒**量级，说明你可能还在家/远程。**必须在公司内网跑**，否则数据没有参考价值。

---

### 第 2 步：在 Windows 上生成探针文件（1 分钟，一次性）

探针必须是**真实规模**的文件。用 13 KB 的小文件测不出任何东西（我第一次就踩了这个坑：Save 只需 50 ms，杀进程窗口根本不存在）。

```powershell
# 建一个专门的探针目录，名字必须含 _probe（脚本会强校验）
New-Item -ItemType Directory -Force '\\你的服务器\共享\_mddt_probe'

# 用真实文件当模板复制过去（不要直接改生产文件）
Copy-Item '\\你的服务器\共享\book.xlsx' '\\你的服务器\共享\_mddt_probe\big_probe.xlsx'
```

**告诉我这个探针文件的真实大小**，我用它校准预期：真实文件 11.44 MB / Save 6.77 s，我这边 13.81 MB / Save 4.2–5.8 s。

> 也可以直接用你真实的 `book.xlsx` 当 `-SourceFile`，脚本只会**复制**它，不改它。
> 探针文件名固定为 `<源文件名>_probe.xlsx`。

---

### 第 3 步：三终端跑 T5（约 10 分钟，需要你在电脑前）

**为什么要三个终端**：Linux 观察者和 Windows 驱动必须**同时**运行，靠 `_phase.txt` 对齐，不依赖时钟同步。

#### 终端 A —— PuTTY：先启动观察者

```bash
./9-linux-observer.sh \
    /你的挂载路径/_mddt_probe/big_probe.xlsx \
    600 200 \
    /你的挂载路径/_mddt_probe/_phase.txt \
    /tmp/t5.csv
```

它立刻开始每 200 ms 采样一次，跑 600 秒。

#### 终端 B —— PuTTY：备用（第 3 阶段的"读者持有"测试）

先放这里，第 3 阶段会用到，暂时不动。

#### 终端 C —— Windows：启动驱动

```powershell
cd G:\projects\mddt_lock_issue
.\collect\9-t3-t5-driver.ps1 `
    -ProbeDir   '\\你的服务器\共享\_mddt_probe' `
    -SourceFile '\\你的服务器\共享\book.xlsx'
```

脚本启动后有 **20 秒倒计时**，让你切到终端 A 确认观察者已经在跑。

#### 中途唯一需要你动手的地方

跑到 `P2_READER_HOLDS` 阶段时，脚本会**等待 30 秒**并在屏幕上提示。这时在**终端 B** 里粘贴：

```bash
exec 9< /你的挂载路径/_mddt_probe/big_probe.xlsx; sleep 60; exec 9<&-
```

这条命令让 Linux **持有**该文件的文件描述符 60 秒，然后 Windows 尝试 Save。
目的是回答：**Linux 正在读的时候，Excel 能不能保存？**

如果这 30 秒你没来得及操作，P2 阶段会作废 —— **不影响其他阶段**，只是少一个数据点。

#### 全部跑完后

**终端 A** 会自动打印判定，形如：

```
distinct md5 values seen (value, count):
  a1b2c3...  412
  d4e5f6...  88

---- VERDICT --------------------------------------
CLEAN: reader only ever saw whole, valid files
```

**把这段判定 + 完整 CSV 贴回来。**

```bash
# 看 CSV（如果很长，只要前面几十行 + 有变化的部分）
cat /tmp/t5.csv | head -50
wc -l /tmp/t5.csv
```

---

## 你要贴回来的东西（就这三样）

1. **预检输出**（第 1 步）
2. **观察者判定段**（第 3 步终端 A 的 `VERDICT` 部分）
3. **CSV**（第 3 步 `cat /tmp/t5.csv`）

如果有异常，再加一条：**Windows 终端 C 的最后 40 行**（脚本会打印 SUMMARY 表）。

---

## 安全说明（为什么不会碰你的生产文件）

| 机制 | 说明 |
|---|---|
| 探针目录名校验 | `ProbeDir` 必须含 `_probe`，否则**硬拒绝**（我实测触发过两次） |
| 目录冲突校验 | `ProbeDir` 不能等于源文件目录，否则**硬拒绝** |
| 源文件只读 | 只 `Copy-Item`，从不打开为可写 |
| 每阶段重建 | 每轮从本地 pristine 副本重新复制探针 |
| 结束清理 | 默认删除探针文件；`-KeepProbe` 才保留 |
| 杀进程对象 | 只杀脚本自己 `Start-ExcelWith` 创建的 Excel PID，不碰你正在用的 Excel |

> ⚠️ 唯一需要你注意：**跑的时候不要有别的 Excel 窗口正在编辑别的文件**。脚本会按"新增 PID"识别自己创建的实例，但如果你正好同时打开 Excel，可能误杀。
> 建议：**先保存并关闭你自己的 Excel**。

---

## 如果中途出问题

| 现象 | 处理 |
|---|---|
| 预检报 `CANNOT SEE THE FILE` | 路径不对，或 Windows 还没生成探针 |
| 预检单次读 > 100 ms | 你可能在 WAN 上，换到公司内网重跑 |
| 观察者报 `bad zip magic` | **这就是结论**，贴回来即可，不要重跑 |
| 驱动报 `REFUSED: ProbeDir must contain _probe` | 目录名改一下，加 `_probe` |
| 驱动卡在 P2 | 等 30 秒自动继续；或在终端 B 执行那条 `exec` 命令 |
| 想跳过杀进程测试 | 加 `-SkipKillTests` |
| 完全没有 Linux 侧条件 | 加 `-SkipReaderHold`，先只跑 Windows 部分（T3） |

---

## 我这边的下一步（不需要你）

| # | 任务 | 状态 |
|---|---|---|
| T19 | 锁的 fencing 回读校验（替代 T2） | 待写 |
| — | `MdlMerge` 骨架 + 单元测试（纯算法，不依赖 Excel） | 待写 |
| — | 仍待你确认的 6 个业务问题 | 见 `collect/7-questions.md` |

**T5 结果回来后**，测试阶段就全部结束了，可以开始写实现代码。