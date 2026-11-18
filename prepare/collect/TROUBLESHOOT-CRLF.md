# 故障：`bad interpreter: No such file or directory`

> 症状：`./9-preflight.sh` 报这个错，但 `ls` 明明能看到文件、`chmod +x` 也做了。

---

## 为什么会这样

这个报错 **99% 不是"文件不存在"**，而是：

> **shebang 那一行末尾带了 `\r`。**

内核读 `/bin/bash\r` → 找不到这个程序 → 报 "No such file or directory"。
**报的是解释器不存在，不是脚本不存在。** 这是最容易误判的地方。

文件从 Windows 通过以下任一方式到达 Linux 时会带上 `\r`：

| 途径 | 是否加 `\r` |
|---|---|
| 复制粘贴到 PuTTY 窗口 | **会**（Windows 剪贴板是 CRLF） |
| WinSCP / FileZilla **Text 模式** | **会** |
| WinSCP / FileZilla **Binary 模式** | 不会 ✅ |
| `scp` / `pscp` | 不会 ✅ |
| 通过 SMB 共享直接 `cp` | 不会 ✅ |
| VS Code 保存时 CRLF | **会** |

我本地生成的这两个文件是**纯 LF、无 BOM**（已验证 178/184 个 LF、0 个 CRLF），
所以问题一定发生在**传输环节**，不是源文件。

---

## 先自己确认（1 条命令）

```bash
file 9-preflight.sh
```

| 输出 | 含义 |
|---|---|
| `Bourne-Again shell script, ASCII text executable` | ✅ 正常 |
| `... with CRLF line terminators` | ❌ **就是这个原因** |
| `ASCII text executable, with very long lines` | 正常 |

或者直接看不可见字符：

```bash
head -1 9-preflight.sh | od -c | head -2
```

正常应该是：
```
0000000   #   !   /   b   i   n   /   b   a   s   h  \n
```
如果看到 `\r \n`，就确认了。

---

## 修法一：一行命令（最快）

```bash
sed -i 's/\r$//' 9-preflight.sh 9-linux-observer.sh
chmod +x 9-preflight.sh 9-linux-observer.sh
file 9-preflight.sh
```

`file` 输出不再有 `CRLF` 就好了。

> `sed -i` 在 `busybox` 上可能不支持，退而用：
> ```bash
> tr -d '\r' < 9-preflight.sh > tmp && mv tmp 9-preflight.sh && chmod +x 9-preflight.sh
> ```

---

## 修法二：用救急包（`sed` 也不好用、或文件已经乱了）

我生成了两个**自解码启动器**，内容全是 base64 —— **任何 `\r` 都不会破坏它**，
因为 base64 在解码前对 shell 来说只是普通文本。

| 文件 | 还原出 |
|---|---|
| `collect/RESCUE-9-preflight.sh` | `9-preflight.sh` |
| `collect/RESCUE-9-linux-observer.sh` | `9-linux-observer.sh` |

用法（在 Linux 上）：

```bash
# 1. 开始接收输入
cat > r.sh
# 2. 把 RESCUE-*.sh 的【全部内容】粘贴进来
# 3. 按 Enter，然后按 Ctrl-D
# 4. 运行
bash r.sh
```

它会**逐字节还原**原文件（我验证过 md5 完全一致）并 `chmod +x`。

> 因为整个 payload 是单行 base64，粘贴时即使行尾被改成 CRLF，
> `base64 -d` 也会忽略 `\r`。**这是最抗传输噪声的方案。**

---

## 修法三：直接用 `bash` 调用，绕过 shebang

如果只是想**先跑起来**看看结果：

```bash
bash 9-preflight.sh /path/to/file.xlsx
```

**注意**：这只绕过 shebang，`\r` 还在，所以脚本内部的字符串比较会带上 `\r`。
例如 `[ "$MAGIC" = "504b0304" ]` 会**永远为假**，导致误报"不是 zip"。
→ **只适合应急，不适合拿数据。**

---

## 用 VS Code 改（三种情况，选你那一栏）

### 情况 A：VS Code 通过 Remote-SSH 直接编辑 Linux 上的文件 ✅ 最常见

这种情况下文件**已经在 Linux 上**，只是内容带 `\r`。

**操作（10 秒）：**

1. 打开 `9-preflight.sh`
2. 看**右下角状态栏**，会显示 `CRLF` 或 `LF`
3. **点它** → 上方弹出选择框 → 选 **`LF`**
4. **`Ctrl+S` 保存**

状态栏从 `CRLF` 变成 `LF` 就成功了。

> 一次只能改一个文件，两个文件各做一次。

**验证：** 在 VS Code 集成终端里跑

```bash
file 9-preflight.sh
```

不再出现 `CRLF line terminators` 即可。

---

### 情况 B：VS Code 在 Windows 本地编辑，再传到 Linux

**先改本地文件的 EOL，再传**：

1. 右下角点 `CRLF` → 选 `LF` → `Ctrl+S`
2. 传到 Linux，**传输模式必须是 Binary**
   - VS Code Remote-SSH：保存即同步，不用管
   - WinSCP / FileZilla：传输设置 → 传输模式 → **二进制**
   - 剪贴板粘贴：**会重新加回 `\r`**，不要用

---

### 情况 C：想一次性防止以后再犯

**（1）把默认 EOL 设为 LF**

`Ctrl+Shift+P` → `Preferences: Open User Settings (JSON)` → 加一行：

```json
{
  "files.eol": "\n"
}
```

> ⚠️ 这只影响**新建文件**，不会回头修已有文件。
> 你现在的两个文件仍需要按情况 A 手动改一次。

**（2）剪贴板粘贴前先确认**

VS Code 里粘贴到 Linux 终端时，如果源是 Windows 剪贴板，
`\r` 仍可能被带入。**优先用 Remote-SSH 直接编辑，不要走剪贴板。**

---

### 情况 D：文件多，不想一个一个点

VS Code 集成终端里一条命令全解决：

```bash
sed -i 's/\r$//' *.sh && chmod +x *.sh && file *.sh
```

---

## 最省事的根治：别用剪贴板传文件

| 方式 | 建议 |
|---|---|
| **WinSCP / FileZilla** | 传输模式设为 **Binary**（默认一般就是） |
| **`pscp` / `scp`** | 首选，绝不改字节 |
| **经 SMB 共享 `cp`** | 最快，文件本来就在共享上 |
| VS Code Remote | 右下角把 CRLF 切成 LF 再保存 |

---

## 验证清单（贴结果前先自查）

```bash
# 1. 行尾干净吗
file 9-preflight.sh            # 不应出现 CRLF
# 2. 语法能过吗
bash -n 9-preflight.sh && echo SYNTAX_OK
# 3. 可执行吗
ls -l 9-preflight.sh           # 应有 x 位
# 4. 能跑吗
./9-preflight.sh --help 2>&1 | head -5
```

四项都过就可以往下走 RUNBOOK 的第 1 步。