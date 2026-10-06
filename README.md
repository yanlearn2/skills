# skills

个人 Codex 技能合集（skill hub）。每个子目录是一个独立 skill，直接放到 `%CODEX_HOME%\skills\`（通常是 `C:\Users\<用户名>\.codex\skills\`）下即可被 Codex 发现。

## 技能列表

| 技能 | 说明 |
| --- | --- |
| [`excel-pwd-opener`](excel-pwd-opener/) | Windows 上接管 `.xlsx/.xls/.xlsm` 关联，双击用已存密码自动打开加密表格；命令行可批量录入、验证、查看记录 |
| [`xlsq`](xlsq/) | 把 `.xlsx` 当表执行 SQL 的 Windows 命令行工具：单文件、离线、能读带打开密码的工作簿，支持工作表正则/合并、正则清洗、导出多种格式与加密报表 |

## 安装某个技能

把对应子目录整个复制到 skills 目录（例如 `C:\Users\<用户名>\.codex\skills\excel-pwd-opener\`）即可；也可以克隆本仓库后：

```powershell
$dest = Join-Path $env:USERPROFILE ".codex\skills"
Copy-Item -Recurse -Force .\excel-pwd-opener $dest
```

## 新增技能

每个技能一个目录，包含 `SKILL.md`（必填，YAML frontmatter 里的 `name` + `description` 决定它何时被选用），可按需附带 `bin/`、`scripts/`、`references/`、`assets/`。
