## 特性

- 转换Excel表格到MarkDown表格
- 转换CSV到MarkDown表格（`-xls file.csv` 或 `-csv file.csv`）
- 支持 `-pretty` 生成列宽对齐、更易读的 Markdown 表格源码
- 支持 `-mmd` 输出带 rowspan/colspan 的 HTML 表格，保留 Excel 合并单元格（适用于 MultiMarkdown）
- 支持Excel单元格带超链接，如果一个单元格，你右键添加了超链接，自动转成`[text](url)`
- 如果Excel里有合并的跨行单元格，在转换后的MarkDown里是分开的单元格，这是因为MarkDown本身不支持跨行单元格
- 如果Excel表格右侧有大量的空列，则会被自动裁剪（前100行采样估宽以保证性能；更后面才出现的列在完整读取时仍会保留）
- 支持指定小数数字的精度
- 支持使用表格首行代替表头（保持空表头）
- 支持指定对齐方式
- 同一个Excel跨表超链接公式，如`HYPERLINK(test_sheet!C9,...)` 会被自动展开成 `[text](url)` 格式
- 同表超链接公式，如`HYPERLINK(C9,...)` 会被自动展开成 `[text](url)` 格式

## 常规用例，文件转换

Mac OS 版本请在命令行下直接使用`exceltk`，不用带exe后缀，MacOS的安装包自动配置好环境变量

- 整个表格
    - `exceltk.exe -t md -xls xxx.xls`
    - `exceltk.exe -t md -xls xxx.xlsx`
    - `exceltk.exe -t md -xls xxx.csv`
    - `exceltk.exe -t md -csv xxx.csv`

- 生成列对齐的美化表格
    - `exceltk.exe -t md -pretty -xls xxx.xlsx`
    - `exceltk.exe -t md -pretty -a c -xls xxx.csv`

- MultiMarkdown 合并单元格（HTML 表格）
    - `exceltk.exe -t md -mmd -xls xxx.xlsx`

- 指定sheet
    - `exceltk.exe -t md -xls xx.xls -sheet sheetname`
    - `exceltk.exe -t md -xls xx.xlsx -sheet sheetname`

- 指定小数数字的精度，例如指定精确到小数点后2位数字
    - `exceltk.exe -t md -p 2 -xls xxx.xls`

## 已移除：剪切板监控（`-t cm`）

- `-t cm`（Windows GUI，监控剪切板并即时转 Markdown）**仅在 0.0.9 提供**，之后版本已移除
- 当前版本请用 `-t md|json|tex` 转换文件
- 如仍需 0.0.9，请到 [Releases](https://github.com/fanfeilong/exceltk/releases) 查找归档资源（不要再依赖 README 里的多版本下载列表）

## 下载

预编译包通过 **GitHub Releases** 发布（打 `v*` 标签后由 Actions 为各 RID 构建自包含包）：

- 最新版：https://github.com/fanfeilong/exceltk/releases/latest
- 全部版本：https://github.com/fanfeilong/exceltk/releases

按系统选择资源：`linux-x64` / `osx-x64` / `osx-arm64`（`.tar.gz`），`win-x86`（`.zip`）。

## 解决在移动设备上表格不能自适应的问题


通过指定`-bhead` 选项解决，使用表格首行代替表头，表头用空的代替：

```
exceltk.exe -t md -bhead -xsl test.xsl
```

输出如下风格的markdown：
```
||||||||||||||
|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|
|**姓名**|**序号**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|
|某某某|34|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|
```

效果如下：

||||||||||||||
|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|:--|
|**姓名**|**序号**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|**积点和**|
|某某某|34|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|6.86|

## 指定对齐方式
```
exceltk -t md -a r -xls example.xlsx
```

`-a` 参数指定对齐方式，可选参数是`l`，`c`，`r`，分别是左对齐、居中对齐、右对齐



# 转换到Json 
  - `exceltk.exe -t json -xls example.xls `

# 转换到TeX
  - `exceltk.exe -t tex -xls example.xls`
  - 使用 `-st n` 拆分表格
  - 使用 `-sn` 把数字拆分，例如`1234656` 会被拆成`1 2 3 4 5 6`, 如果表太大时有用
# 在 Linux 上构建与使用

需要安装 [.NET 10 SDK](https://learn.microsoft.com/dotnet/core/install/linux)（LTS）。仓库根目录 `global.json` 会约束使用 .NET 10 SDK。

```bash
dotnet build src/Exceltk/Exceltk.csproj -c Release
dotnet run --project src/Exceltk/Exceltk.csproj -c Release -- -t md -xls src/test/test1.xlsx

dotnet publish -r linux-x64 src/Exceltk/Exceltk.csproj -c Release
./src/bin/net10.0/linux-x64/publish/exceltk -t md -xls src/test/test1.xlsx
```

说明：`dotnet run` 传参时请在参数前加 `--`，避免 `-t` / `-a` 被 dotnet 自己吃掉。CI 见 `.github/workflows/ci.yml`。更多平台说明见英文 README 的 How to build。
