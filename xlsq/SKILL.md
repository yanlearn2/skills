---
name: xlsq
description: Query and export xlsx workbooks as SQL tables with the xlsq CLI, including password-protected files, sheet-name regex selection, and regex data cleaning. Use when a task needs to read, filter, join, or export data from .xlsx files from the command line; not for building spreadsheet applications or reading legacy .xls.
---

# xlsq

把 xlsx 当表执行 SQL 的 Windows 命令行工具：单文件、离线、能读**带打开密码**的工作簿。
表名固定为 `data`，工作表用参数选（不要把表名拼进 SQL）。

## 开始前

本 skill 自带可执行文件 `<skill目录>\bin\xlsq.exe`（约 6.3MB，静态链接，不依赖 VC++ 运行库），
拿到 skill 就能用，**不要假设目标机器装过 xlsq**。

先把自带的 exe 解析成绝对路径再调用，PowerShell 例：

    $xlsq = (Resolve-Path "<本 SKILL.md 所在目录>\bin\xlsq.exe").Path
    & $xlsq --version
    & $xlsq sheets 你的文件.xlsx --detail
    & $xlsq -f 你的文件.xlsx --sheet Sheet1 -H --out-header "SELECT * FROM data LIMIT 10"

常见安装位置是 `%CODEX_HOME%\skills\xlsq\bin\xlsq.exe`（通常是
`C:\Users\<用户名>\.codex\skills\xlsq\bin\xlsq.exe`）。PATH 里已能直接跑 `xlsq` 时用 PATH 里的也行；
否则直接用完整路径，不要在 PATH 上折腾。

**参数不要凭记忆猜**：`xlsq --help` 是权威的完整参数表。

三步例行检查，能省掉绝大多数来回：

1. `xlsq sheets <文件> --detail` —— 列出工作表名、行列数、非空单元格数。
   工作表名可能带**尾部空格**（例如 `103钴盐车间 `）或括号，照抄名字用 `--sheet`。
   空表会显示 `读取失败: 表格没有数据` 但仍以 0 退出，不影响其它工作表，不用当报错处理。
2. `xlsq -f 文件 --sheet 名称 -H --dry-run "SELECT 1"` —— 先看会解析出哪些列名。
3. 结果不符合预期时加 `--stats` —— 打印打开、读取导入、SQL 各阶段耗时与缓存命中情况。

## 选工作表与读取参数

- `--sheet <名称>`：精确匹配，含空格与括号的名字可直接用。
- `--sheet-regex <RE2>`：正则批量选。选到多张表时必须加 `--merge`（按列名对齐合并、缺列补 NULL），
  否则会明确报错，而不是悄悄只读一张。
- `-H/--header`：首行作表头。空列名与重名列会自动替换成单元格名（如 `B1`），
  因此重名表头不会导致建表失败。
- `-s/--skip <N>`：跳过前 N 行；`--cell <如 C1>`：从某单元格起读（一张表上并排多张表时用）。
  与 `-H` 一起用时表头取第 N+1 行，且**读出的列名与不带 `-s` 时一致**（H 与 skip 在内部叠加）；
  N 大到把数据行也跳过会报「表格没有数据」并退出 3，而不是返回空结果。
- `--columns a,b,c`：**显式**只读这几列，用于宽表只取几列的场景。用了它之后
  SQL 里引用其它列会报错并提示——这是有意的，不要用它去猜列。
- `--empty-row-limit <N>`：连续空行达到 N 行即停止读取，默认 1000，`0` 表示不限制。
  有些文件声明的已使用范围虚高（到上百万行）；当 `sheets --detail` 里的行数远大于
  非空单元格数时，保持默认或调低，否则会把大量空行灌进内存与数据库。

## 密码

优先 `--password-file <文件>`（读首行）或环境变量 `XLSQ_PASSWORD`；也可用
`--password-stdin`、`--ask-password`。**避免 `-p`**：明文密码会出现在进程命令行里
（任务管理器、各类日志都能读到），程序会对此打警告。支持 ECMA-376 Standard（AES-128+SHA-1）
与 Agile（AES-256+SHA-512）。解密在内存中完成，不落明文临时文件。

## SQL 里的自定义函数

- `regexp_like(文本, 正则[, 'i'|'m'])` —— 是否匹配；`i` 忽略大小写，`m` 多行。
- `regexp_extract(文本, 正则[, 分组号])` —— 提取分组，默认 0；无匹配返回空串。
- `regexp_replace(文本, 正则, 替换文本)` —— 批量清洗。
- `excel_date(序列号)` —— Excel 序列号转 `YYYY-MM-DD`（基准 1899-12-30），
  配合 `date('now','-1 month')` 做区间筛选。

**`excel_date` 碰到非数字文本会让整条查询报错退出 2**，不是跳过该行。日期列里混有
「标准」这类模板文本很常见（车间表就有），合并多表后容易踩到——只要那条 SQL 扫过
文本行就会失败：

    excel_date(送样日期)                                 -- 356 行文本 -> 整条查询失败，退出码 2
    WHERE regexp_like(送样日期,'^[0-9]+(\.[0-9]+)?$')   -- 先挡掉非数字行，可用

两个相邻的坑：`WHERE typeof(送样日期) IN ('integer','real')` 会**静默返回 0 行**——
导入后日期列一律是 text 类型；而 `excel_date(CAST(送样日期 AS REAL))` 又能算出值，
容易误判。日期的稳妥写法是先用上面的正则挡掉非数字行，再调 `excel_date`。

正则是 RE2 语义，没有回溯、不会因正则写错而卡死。**不支持的高级语法（零宽断言
`(?=…)`/`(?!…)`/`(?<=…)`、反向引用等）是静默不匹配、返回 0，不报错**——例如
`regexp_like(样品名称,'(?=Co)')` 得到 0，看着像「没数据」，其实是语法不被支持。
拿不准时用 `regexp_like(值,'Co')` 这类能直观看出的模式对照验证一次（`'CoCl2前液槽'` 上用 `(?=Co)` 得 0、用 `'Co'` 得 1）。

**空单元格导入为 NULL，而不是空字符串**，所以筛掉空值要写 `列 IS NOT NULL`。

## 输出

`-o` 选格式（CSV、TSV、LTSV、JSON、JSONL、YAML、TBLN、MD、AT、VF、RAW、XLSX），
`-O` 指定文件名，`--out-header` 输出表头。

导出 xlsx：`-o XLSX -O 报表.xlsx`，可用 `--out-sheet`、`--out-cell`、`--out-header`；
**目标文件已存在时必须加 `--clear-sheet`**——本实现整文件写出，会覆盖其中其它工作表，
因此要求显式确认。导出加密报表用 `--out-password`。

多语句可用（先 `UPDATE` 再 `SELECT`），输出最后一条产生结果集的语句。

## 缓存与配置

缓存默认开启，键包含文件路径、大小、mtime、密码指纹**和读取参数**，任一变化即失效；
同一列集重复查询会直接复用，跳过解密与解析。相关操作：`--no-cache`、`--refresh-cache`、
`cache list|prune|clear`；`cache info <键>` 必须带上键（键从 `cache list` 取），
只写 `xlsq cache info` 会以用法错误退出 1。缓存文件整体加密，密钥在缓存目录的
`.masterkey`——共享机器上建议 `--no-cache`。

配置分层：命令行 > 环境变量 > `./xlsq.toml` > `%APPDATA%\xlsq\config.toml`。
`xlsq config print` 会逐项显示生效值**及其来源**，排查「为什么这个值生效了」时先看它。

## 退出码（脚本分流）

`0` 成功；`1` 参数或用法问题（含工作表不存在 / 未指定）；`2` SQL 错误；`3` 文件、密码或数据问题。

实测补充（脚本按码分流时按这几条更准）：

- 命令行解析失败（未知参数、缺值）也返回 `2`，不只是 SQL 错误。
- `--columns` 声明的列不存在 → 退出 `3`（提示「列不存在」）；SQL 里引用不存在列才走 `2`。
- 因 `--skip` 跳空全部行而报「表格没有数据」→ 退出 `3`。

## 已知坑

- SQLite 的列名比较**不区分大小写**：表头同时出现 `si` 与 `Si` 会被视为重名
  （工具已自动去重；但你自己在 SQL 里拼列名时要注意）。
- `--columns` 与 `--merge` 不能同时使用（合并后的列集由各表表头决定）。
- 用 PowerShell 脚本包裹调用时，脚本文件要存成 **UTF-8 带 BOM**，否则 PS 5.1 会把中文 SQL
  读成乱码；`.bat` 则不能带 BOM。
- 路径含空格、括号、`&` 时把整个参数用引号包住。

## 典型调用

    # 加密文件 + 多车间合并 + 正则筛选 + 导出加密报表
    xlsq.exe -f 中控样品.xlsx --password-file keys.txt --sheet-regex '^10[12]车间' --merge -H --out-header -o XLSX -O 汇总.xlsx --out-password "导出密码" "SELECT excel_date(送样日期) AS 日期, 样品名称, Co FROM data WHERE regexp_like(送样日期,'^[0-9]+(\.[0-9]+)?$') AND regexp_like(样品名称,'^一浸精滤') AND Co IS NOT NULL"
