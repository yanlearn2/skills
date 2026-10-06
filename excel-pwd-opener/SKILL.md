---
name: excel-pwd-opener
description: 用 ExcelAutoOpener（opener.exe / opener-cli.exe）在 Windows 上接管 .xlsx/.xls/.xlsm 关联，双击就用已存密码自动打开加密表格，并可用命令行批量录入、验证、查看记录。当任务涉及“打开带密码的 Excel / 记住表格密码 / 注册文件关联 / 批量验证密码 / 在脚本或 Agent 里打开加密工作簿”时使用。
metadata:
  short-description: 加密 Excel 自动填密码打开（GUI + CLI）
---

# excel-pwd-opener

Windows 小工具 **ExcelAutoOpener**：接管 `.xlsx/.xls/.xlsm` 的打开方式，**双击就用事先存好的密码自动打开**
加密表格，不必每次手输密码；也能被脚本 / RPA / AI Agent 用命令行调用。

本 skill 自带可执行文件 `<本 SKILL.md 所在目录>\bin\opener.exe` 与 `...\bin\opener-cli.exe`，
拿到 skill 就能用，**不要假设目标机器装过它**。先解析成绝对路径再调用：

```powershell
$bin = (Resolve-Path "<本 SKILL.md 所在目录>\bin").Path
& "$bin\opener.exe" --version
```

**参数不要凭记忆猜**：`& "$bin\opener-cli.exe" -h` 是权威的完整命令表（`--json -h` 会返回机器可读的能力清单）。

## 选哪个 exe（重要）

| 文件 | 子系统 | 用途 |
| --- | --- | --- |
| `opener-cli.exe` | 控制台 | **脚本 / 自动化 / Agent 首选**：cmd 退出码可靠、PowerShell 重定向正常、无窗口闪动 |
| `opener.exe` | GUI | 双击打开文件、打开设置窗口；也能在控制台当命令行用，但被 Windows PowerShell 5.1 的 `>` 重定向时会写出空文件 |

两者**参数、输出、退出码含义完全一致**（共用同一套命令实现），差别只在运行形态。
写脚本、做自动化、让 Agent 调用 → 一律用 `opener-cli.exe`，并加 `--json`。

## 安装 / 卸载

接管文件关联只要当前用户权限，不弹 UAC、不联网：

```powershell
& "$bin\opener.exe" --install      # 接管 .xlsx/.xls/.xlsm、注册右键动词、建开始菜单
& "$bin\opener.exe" --uninstall    # 还原关联、清理快捷方式
```

安装会把 `opener.exe` 和 `opener-cli.exe` 复制到 `%LOCALAPPDATA%\ExcelAutoOpener\`。

- 只想接管 / 还原关联：`--assoc` / `--restore`
- 密码库存于 `%APPDATA%\ExcelAutoOpener\store.json`（按当前用户用 Windows DPAPI 加密），配置在同目录 `config.json`

## 打开加密表格（进入“编辑模式”的三种方式）

- 右键 → **编辑**
- **Shift + 双击**
- **按住 Ctrl 再双击**

> Ctrl 的判定发生在**按下的那一刻**：先按住 Ctrl 不松、再双击；先双击再补按 Ctrl 无效。

编辑模式要求记录里存了**编辑密码**（修改权限密码）。存对了就直接以可写方式打开，
不再弹密码框；没存或存错才会让你手输。

## 命令行速查

| 命令 | 作用 |
| --- | --- |
| `--add <路径> <打开密码> [编辑密码]` | 录入 / 覆盖一条记录；`-` 表示该把密码留空 |
| `--remove <路径>`（别名 `--delete`） | 删除记录（**不动磁盘上的表格文件**），幂等 |
| `--list` | 列出密码库里所有记录 |
| `--test <路径>` | 用库里存的密码打开验证一次 |
| `--test-all` | 逐条验证所有记录 |
| `--engine [wps\|excel]` | 查看或设置使用的办公软件 |
| `--install` / `--uninstall` / `--assoc` / `--restore` | 安装 / 卸载 / 只接管 / 只还原关联 |
| `--edit <路径>` / `--readonly <路径>` | 显式以编辑 / 只读方式打开 |
| `<路径>` | 按记录设置打开（被文件关联调用） |

全局选项：`--json`（等价 `-j`，可放在任意位置）机读输出；`-h/--help`；`-V/--version`。

## 退出码（脚本分流）

`0` 成功；`1` 运行错误（文件不存在、I/O、环境问题）；`2` 用法错误，或 `--test/--test-all` 判定密码不正确。
`--remove` 对“记录本就不存在”返回 `0`（幂等），方便脚本重复执行。

## 机读输出（`--json`）

```powershell
& "$bin\opener-cli.exe" --json --list
# {"ok":true,"count":1,"entries":[{"path":"D:\\a.xlsx","has_open_password":true,"has_edit_password":false,"open_editable":false,"stored_passwords":"open"}]}
```

关键字段：`--add/--remove` → `action`、`removed`；`--list` → `count`、`entries`；
`--test` → `result`（`ok | bad_open_password | unverified_edit | unavailable`）；
`--test-all` → `total_tested`、`passed`、`failed`、`unverified`、`bad_paths`。

## 已知坑

- **只有真打开表格才需要 Office**：`--test` / `--test-all` 会真启动 WPS / Excel（约 1~4 秒）；
  `--add` / `--list` / `--remove` / `--install` / `--engine` 只读写本地记录，毫秒级、不启动办公软件。
- **批量录入就别逐条测**：先 `--add` 批量写库，再单独 `--test-all` 复查。一条条 `--test` 会反复启停 Office，很慢。
- **PowerShell 5.1 的 `>` 重定向**：对 `opener.exe`（GUI 子系统）会写出空文件；改用 `opener-cli.exe`，或写成 `... | Out-File`。
- **密码会出现在命令行里**：`--add` 的密码明文可见于进程列表 / 脚本历史。敏感场景从环境变量或凭据库读取，或在 GUI 设置窗口手工录入。
- **换账户 / 换机器读不到密码**：DPAPI 按当前用户加密，换 Windows 用户或重装系统需重录。
- **写含中文的 `.ps1` 要存成 UTF-8 带 BOM**（PowerShell 5.1），`.bat` 则不能带 BOM。
- `--probe` 是内部动词（设置窗口后台调用），脚本不要依赖。

## 参考

- 面向普通用户的图文说明：[references/使用指南.md](references/使用指南.md)
- 面向自动化的完整 CLI 手册（命令、退出码、JSON 形状、示例脚本、Agent 接入建议）：[references/CLI使用指南.md](references/CLI使用指南.md)
