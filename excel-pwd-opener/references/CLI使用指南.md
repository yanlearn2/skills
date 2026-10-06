# opener 命令行使用指南（面向自动化 / 脚本 / AI Agent）

本指南假定你通过脚本、任务计划、RPA 或 AI Agent 调用 `opener` 来处理加密表格。
如果你只想「双击自动打开」，请看 `使用指南.md`，不用关心命令行。

---

## 0. 先选对可执行文件（很重要）

发行包里有**两个**可执行文件，命令完全一样，区别只在「运行形态」：

| 文件 | 子系统 | 用途 |
| --- | --- | --- |
| `opener.exe` | Windows GUI | 双击打开文件、打开设置窗口；也可当命令行用 |
| `opener-cli.exe` | 控制台 | **脚本/自动化首选**：cmd 里退出码可靠、PowerShell 重定向正常、不弹任何窗口 |

什么时候用 `opener-cli.exe`：

- 在 cmd / .bat 里要拿准确的 `%errorlevel%`；
- 在 PowerShell 里要 `opener-cli --list > out.txt` 这种重定向（GUI 版的 `>` 在 Windows PowerShell 5.1 会写出空文件）；
- 在 RPA / Agent 里既要**读输出**又要**读退出码**，还不想有任何窗口闪动。

> 两者输出内容、参数、退出码含义完全一致；`opener-cli.exe` 只是更「像个正常命令行程序」。

---

## 1. 装上并验证

```bat
opener.exe --install        :: 接管 .xlsx/.xls/.xlsm、建开始菜单（一条命令搞定，无需管理员）
opener-cli.exe --version    :: 打印版本
opener-cli.exe -h           :: 打印完整帮助（简体中文，已修复乱码）
```

安装是「自助式」：程序把自己复制到 `%LOCALAPPDATA%\ExcelAutoOpener\`，并注册文件关联。
卸载：

```bat
opener.exe --uninstall      :: 还原关联、清理快捷方式（等价于 --restore 全套）
```

---

## 2. 命令速查

| 命令 | 作用 |
| --- | --- |
| `opener --add <路径> <打开密码> [编辑密码]` | 录入/覆盖一条记录；`-` 表示该把密码留空 |
| `opener --remove <路径>`（别名 `--delete`） | 删除记录（**不动磁盘上的表格文件**），幂等 |
| `opener --list` | 列出密码库里所有记录 |
| `opener --test <路径>` | 用库里存的密码验证一次（**默认离线核对，不启动 Excel/WPS**） |
| `opener --test-all` | 逐条验证库里所有记录（同样优先离线核对，因此很快） |
| `opener --engine [wps\|excel]` | 查看或设置使用的办公软件 |
| `opener --install` / `--uninstall` | 安装 / 卸载 |
| `opener --assoc` / `--restore` | 只接管 / 只还原文件关联 |
| `opener --edit <路径>` / `--readonly <路径>` | 显式以编辑 / 只读方式打开 |
| `opener <路径>` | 按记录设置打开（被文件关联调用） |

全局选项：`--json`（等价 `-j`）机读输出；`-h/--help`；`-V/--version`。
`--json` 可放在任意位置：`opener --json --list` 与 `opener --list --json` 等价。

### 密码怎么传

- 打开密码、编辑密码各一把，按位置传：`--add <路径> <打开密码> <编辑密码>`。
- 只有一把（比如只有打开密码）：`opener --add "D:\a.xlsx" "P@ssw0rd"`。
- 只填编辑密码、打开密码留空：`opener --add "D:\a.xlsx" - "P@ssw0rd"`。
- ⚠️ 密码会出现在命令行 / 进程列表里。敏感场景见 [第 6 节](#6-安全与权限)。

---

## 3. 退出码（脚本判断的核心）

| 退出码 | 含义 |
| --- | --- |
| `0` | 成功；`--test` 表示密码验证通过 |
| `1` | 运行错误：文件不存在、I/O 失败、环境问题、未能确认 |
| `2` | 用法错误；或 `--test`/`--test-all` 发现**密码不正确** |

> 从 0.1.2 起，`--test` **不要求能打开文件，只核对密码对不对**：绝大多数情况直接
> 解析文件内的密码校验信息（毫秒级、不启动办公软件）。只有「编辑密码藏在加密内容里」
> 这种表才会短暂启动一次套件补验。所以「打不开」不再是判定密码错误的前提。

`--remove` 对「记录本就不存在」返回 `0`（幂等），方便脚本重复执行。

**cmd / .bat：**

```bat
opener-cli.exe --test "D:\报表\2026Q1.xlsx"
if errorlevel 2 ( echo 密码不对 ) else if errorlevel 1 ( echo 运行出错 ) else ( echo OK )
```

**PowerShell：**

```powershell
opener-cli.exe --test "D:\报表\2026Q1.xlsx"
switch ($LASTEXITCODE) {
  0 { "密码正确" }
  2 { "密码不正确" }
  default { "未能确认 / 运行出错" }
}
```

---

## 4. 机读输出（`--json`）

加 `--json`（或设环境变量 `OUTPUT_FORMAT=json` / `OPENER_FORMAT=json`），结果变成**单行 JSON**，
机读时不会有任何装饰性文字，退出码仍按第 3 节规则返回。

```bat
opener-cli.exe --json --list
:: {"ok":true,"count":1,"entries":[{"path":"D:\\a.xlsx","has_open_password":true,"has_edit_password":false,"open_editable":false,"stored_passwords":"open"}]}
```

各命令的 JSON 形状：

| 命令 | JSON 关键字段 |
| --- | --- |
| `--add` / `--remove` | `{"ok":true,"action":"add\|remove","path":"...","removed":true\|false}` |
| `--list` | `{"ok":true,"count":N,"entries":[{"path","has_open_password","has_edit_password","open_editable","stored_passwords"}]}` |
| `--test` | `{"ok":bool,"path":"...","result":"ok\|open_password_ok\|bad_open_password\|bad_edit_password\|unavailable","message":"..."}` |
| `--test-all` | `{"ok":bool,"total_tested","passed","failed","unverified","bad_paths":[...],"entries":[{"path","result"}]}` |

`result` 取值含义：`ok` 打开密码与编辑密码都核对通过；`open_password_ok` 打开密码通过、
编辑密码这次没验（不算错）；`bad_open_password` / `bad_edit_password` 对应那把密码不正确；
`unavailable` 环境原因未能验证。
| `--version` / `--help` | `{"ok":true,"name":"opener","version":"...", ...}` |

在脚本里解析（PowerShell）：

```powershell
$r = opener-cli.exe --json --list | ConvertFrom-Json
foreach ($e in $r.entries) {
  "{0}  打开密码:{1}  编辑密码:{2}" -f $e.path, $e.has_open_password, $e.has_edit_password
}
```

在 Python 里解析：

```python
import json, subprocess
p = subprocess.run([r"C:\path\opener-cli.exe", "--json", "--list"],
                   capture_output=True)
p.check_returncode()
data = json.loads(p.stdout.decode("utf-8"))   # 输出恒为 UTF-8
print(data["count"], "条记录")
```

---

## 5. 典型自动化脚本

### 5.1 批量录入一批已知密码的表格

```powershell
$csv = Import-Csv .\passwords.csv   # 列：Path,OpenPwd,EditPwd
foreach ($row in $csv) {
  $edit = if ($row.EditPwd) { $row.EditPwd } else { "-" }
  & opener-cli.exe --add $row.Path $row.OpenPwd $edit
  if ($LASTEXITCODE -ne 0) { Write-Warning "录入失败: $($row.Path)" }
}
& opener-cli.exe --list              # 核对
```

> 批量录入不必逐条测试：`--add` 是纯写库、毫秒级、不打开表格。想复查再单独跑 `--test-all`。

### 5.2 每日巡检：发现密码失效的表格

```bat
opener-cli.exe --json --test-all > "%TEMP%\pwd-check.json"
if errorlevel 2 (
  echo 有密码不正确，请查看 %TEMP%\pwd-check.json
  exit /b 1
)
echo 全部正常
```

### 5.3 清理某个表格的记录后再重录

```powershell
& opener-cli.exe --remove "D:\数据\月度销售.xlsx"
& opener-cli.exe --add    "D:\数据\月度销售.xlsx" (Read-Host "打开密码")
& opener-cli.exe --test   "D:\数据\月度销售.xlsx"
```

---

## 6. 安全与权限

- **密码可见性**：`--add` 的密码在命令行里明文出现，会被进程列表、脚本历史、日志看到。
  敏感场景建议：脚本里从环境变量 / 凭据库读取后传入，避免写死在脚本里；
  或在 GUI 设置窗口里手工录入（不经过命令行）。
- **存储**：密码经 Windows DPAPI 按**当前用户**加密后存于
  `%APPDATA%\ExcelAutoOpener\store.json`；只有同一 Windows 账户能解密。
  换账户 / 换机器读不到，需要重录。
- **权限**：所有文件只写当前用户目录与 HKCU，不需要管理员权限，安装不弹 UAC，且**不联网**。
- **`--probe` 是内部动词**（供设置窗口后台调用），不建议脚本直接依赖。

---

## 7. 环境与兼容性

- 只有「真正打开表格」才需要已安装 **WPS 表格** 或 **Microsoft Excel**：
  `--test` / `--test-all` 会真实打开一次表格来验证（会启动办公软件，约 1～4 秒）；
  `--add` / `--list` / `--remove` / `--install` / `--engine` 只读写本地记录，不会启动办公软件。
- **编辑密码的验证**：`--test` 会真的用「可写方式」打开一次表格来核对编辑密码；
  只有能可写打开才算通过。若返回 `unverified_edit`，通常是文件正被占用、或云盘里的
  文件还没下载完（不代表密码一定错），换台机器或等文件就绪后重测。
- 输出编码：写真实控制台用 Unicode（不再乱码）；重定向到文件 / 管道时恒为 **UTF-8**。
- `opener.exe`（GUI 版）在 Windows PowerShell 5.1 里用 `>` 重定向可能得到空文件——
  这是 GUI 子系统程序的通病，换成 `opener-cli.exe` 或 `... | Out-File` 即可。

---

## 8. 给 AI Agent 的接入建议

- **调用 `opener-cli.exe`**，并始终加 `--json`：输出是稳定的单行 JSON，无需解析中文。
- 用**退出码**决定流程：`0` 成功、`1` 环境/运行错误（可重试）、`2` 密码或用法问题。
- 建议把命令白名单限定为：
  `--list`、`--add`、`--remove`、`--test`、`--test-all`、`--engine`、`--version`、`--help`。
  避免让 Agent 调用 `--install` / `--uninstall` / `--probe`（会改变系统状态或属于内部用法）。
- `--help` 在 `--json` 下会额外返回机器可读的 `commands` 数组，可用于让 Agent 自描述能力。
