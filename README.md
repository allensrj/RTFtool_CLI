# RTF Tools CLI v0.4 使用说明

> **前置条件：** Windows 系统，已安装 Microsoft Word。  
> 除 `combine` 命令外，其余命令均通过 Word COM 接口操作，运行前请**关闭所有已打开的 Word 文档**。

---

## 总览

```
rtftool <command> [flags]
```

| 命令 | 功能 |
|------|------|
| `check` | RTF 页码校验 |
| `combine` | RTF 合并（Specify 样式，支持 TOC 和页码刷新） |
| `combine-docx` | DOCX/RTF 通用合并（General 样式，支持可选 TOC） |
| `convert` | RTF → PDF / DOCX 转换 |
| `docx2rtf` | DOCX → RTF 转换（支持单文件和批量） |

查看帮助：

```bash
rtftool --help
rtftool <command> --help
```

---
## 使用步骤
1. 去 releases 里下载 [rtftool_v0.4.0_windows_amd64.zip](https://github.com/allensrj/RTFtool_CLI/releases/download/v0.4.0/rtftool_v0.4.0_windows_amd64.zip)
2. 解压 rtftool_v0.4.0_windows_amd64.zip，将其中的 rtftool.exe 放置到指定位置或加入环境变量均可。PDF 优化器已内嵌于 rtftool.exe；
3. 如果是指定位置，就这样运行 C:\Projects\rtftool.exe --help

## 1. check — RTF 页码校验

校验指定文件夹（含子文件夹）内所有 RTF 文件的页码一致性。通过 Word COM 获取实际渲染页数，同时从 RTF 文本中解析 `Page 1 of N` 标记，对比两者。不一致时自动定位到出错的节。

### 参数

| 参数 | 短写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--dir` | `-d` | string | 是 | 要校验的文件夹路径（递归扫描子文件夹） |

### 用法示例

```bash
# 校验单个文件夹
rtftool check -d "C:\Projects\RTF_Output"

# 使用完整参数名
rtftool check --dir "D:\reports\study_123\rtf_files"
```

### 输出示例

```
[INFO] Initializing environment...
[INFO] No WINWORD.EXE process found.
[Worker 0] Processing [1/10]: T-14-1-1.rtf
[Worker 1] Processing [2/10]: T-14-2-1.rtf
...

========================================
Total: 10 | Matched: 9 | Mismatched: 1 | Failed: 0
Duration: 25.3s
```

### 退出码

| 退出码 | 含义 |
|--------|------|
| 0 | 所有文件页码一致 |
| 1 | 运行出错（路径不存在等） |
| 2 | 存在页码不一致的文件 |

---

## 2. combine — RTF 合并（Specify 样式）

将多个 RTF 文件按指定顺序合并为一个 RTF 文件，支持自动生成目录（TOC）和全局页码刷新。此命令**不依赖 Word COM**，纯文本操作，速度快。

### 参数

| 参数 | 短写 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|------|--------|------|
| `--out-dir` | `-o` | string | 是 | — | 输出目录 |
| `--out-name` | `-n` | string | 是 | — | 输出文件名（不含 `.rtf` 扩展名） |
| `--toc` | `-t` | bool | 否 | false | 是否在合并文件头部插入目录 |
| `--toc-rows` | — | int | 否 | 23 | 目录每页显示行数 |
| `--refresh-page` | `-p` | bool | 否 | false | 是否重写所有 `Page X of Y` 页码 |
| `--input-dir` | `-i` | string | 否 | — | 从指定文件夹自动扫描 RTF 文件（替代手动列出文件） |
| `--sort` | `-s` | bool | 否 | true | 是否按 T/F/L 前缀和数字自动排序 |

### 文件来源

有两种方式指定要合并的文件，**任选其一**：

1. **位置参数**：在命令后直接列出所有 RTF 文件路径，按列出顺序合并
2. **`--input-dir`**：指定一个文件夹，自动扫描其中所有 `.rtf` 文件

### 自动排序规则（`--sort`）

启用排序时，文件按以下规则排列：
- 以 `T` 开头的文件排最前（Table）
- 以 `F` 开头的文件排中间（Figure）
- 以 `L` 开头的文件排最后（Listing）
- 其他文件排末尾
- 同类文件按文件名中的数字部分升序排列（如 `T-14-1` < `T-14-2` < `T-14-10`）

### 用法示例

```bash
# 方式1：手动指定文件和顺序
rtftool combine -o "C:\Output" -n "Combined_Report" -t -p ^
  "C:\RTF\T-14-1.rtf" "C:\RTF\T-14-2.rtf" "C:\RTF\F-14-1.rtf"

# 方式2：扫描文件夹，自动排序，生成目录并刷新页码
rtftool combine -i "C:\RTF\source" -o "C:\Output" -n "Combined_Report" --toc --refresh-page

# 方式3：扫描文件夹，不排序（按文件系统顺序），不加目录
rtftool combine -i "C:\RTF\source" -o "C:\Output" -n "Combined_Report" --sort=false

# 自定义目录每页行数
rtftool combine -i "C:\RTF\source" -o "C:\Output" -n "Report" -t --toc-rows 30
```

### RTF 输入文件要求

为确保合并和目录生成正确，输入的 RTF 文件需满足：

1. **标题格式**：必须存在 IDX 书签标记，标题格式为 `\s999 \b [标题文本] \b0`
2. **页码格式**：支持 `Page 1 of 5` 或 `Page 1 / 5` 格式
3. **文档边界控制符**：至少包含 `\widowctrl`、`\sectd` 或 `\info` 之一
4. **页面尺寸**：第一个 RTF 文件头部需包含 `\pgwsxn`（宽）和 `\pghsxn`（高）标签

---

## 3. combine-docx — DOCX/RTF 通用合并（General 样式）

通过 Word COM 将多个 DOCX 或 RTF 文件按顺序合并为一个 DOCX 文件。支持可选目录（TOC）插入。

### 参数

| 参数 | 短写 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|------|--------|------|
| `--out-dir` | `-o` | string | 是 | — | 输出目录 |
| `--out-name` | `-n` | string | 是 | — | 输出文件名（不含扩展名） |
| `--input-dir` | `-i` | string | 否 | — | 从指定文件夹扫描 `.docx` 和 `.rtf` 文件 |
| `--sort` | `-s` | bool | 否 | true | 是否按 T/F/L 前缀和数字自动排序 |
| `--toc` | `-t` | bool | 否 | false | 在输出 DOCX 开头插入目录 |

### 文件来源

与 `combine` 命令相同，支持两种方式：
1. **位置参数**：直接列出文件路径
2. **`--input-dir`**：自动扫描文件夹

### 用法示例

```bash
# 手动指定文件
rtftool combine-docx -o "C:\Output" -n "Merged_Report" ^
  "C:\docs\part1.docx" "C:\docs\part2.docx" "C:\docs\part3.rtf"

# 扫描文件夹自动排序
rtftool combine-docx -i "C:\docs\source" -o "C:\Output" -n "Merged_Report"

# 扫描文件夹自动排序，并插入目录
rtftool combine-docx -i "C:\docs\source" -o "C:\Output" -n "Merged_Report" --toc

# 扫描文件夹，不排序
rtftool combine-docx -i "C:\docs\source" -o "C:\Output" -n "Merged_Report" --sort=false
```

### 注意事项

- 运行前请关闭所有 Word 文档
- 合并过程中会产生一个中间 RTF 文件（用于处理 `\pgnrestart`），最终输出为 DOCX
- 程序会自动终止残留的 WINWORD.EXE 进程

---

## 4. convert — RTF → PDF / DOCX 转换

将单个 RTF 文件转换为 PDF 和/或 DOCX 格式。PDF 输出会自动调用内嵌的优化器进行优化（书签展开 + Fast Web View）。

### 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `<source.rtf>` | 位置参数 | 是 | 源 RTF 文件路径 |
| `--pdf` | bool | 否* | 转换为 PDF |
| `--docx` | bool | 否* | 转换为 DOCX |

> \* `--pdf` 和 `--docx` 至少需要指定一个。

### 输出文件位置

转换后的文件会生成在源 RTF 文件的**同目录**下，文件名与源文件相同（扩展名不同）。

### 用法示例

```bash
# 同时转换为 PDF 和 DOCX
rtftool convert --pdf --docx "C:\reports\report.rtf"

# 只转换为 PDF
rtftool convert --pdf "D:\output\combined.rtf"

# 只转换为 DOCX
rtftool convert --docx "C:\reports\annual_report.rtf"
```

### 注意事项

- 运行前请关闭所有 Word 文档
- 转换时间取决于文件大小，大文件（如 20MB RTF）可能需要约 10 分钟
- PDF 转换完成后会自动调用优化器（`optimize_pdf.exe`，已内嵌于程序中）
- 转换过程中会修改文档属性：标题清空、作者设为 `ZaiLab`

---

## 5. docx2rtf — DOCX → RTF 转换

将单个 DOCX 文件或整个文件夹中的所有 DOCX 文件批量转换为 RTF 格式。已存在同名 RTF 文件的会自动跳过。

### 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `<path>` | 位置参数 | 是 | 单个 `.docx` 文件路径，或包含 `.docx` 文件的文件夹路径 |

### 输出文件位置

RTF 文件生成在源 DOCX 文件的**同目录**下，文件名相同（扩展名为 `.rtf`）。

### 用法示例

```bash
# 转换单个文件
rtftool docx2rtf "C:\docs\report.docx"

# 批量转换整个文件夹
rtftool docx2rtf "C:\docs\batch_folder"
```

### 输出示例

```
[INFO] Target path: C:\docs\batch_folder
----------------------------------------
[INFO] Scanning target path...
[INFO] Found 5 file(s) to convert.
[INFO] Cleaning up background Word processes...
[INFO] Initializing Word COM components...
  [1/5] converting: report1.docx -> OK
  [2/5] converting: report2.docx -> OK
  [3/5] converting: report3.docx -> Skipped (identical RTF already exists).
  [4/5] converting: report4.docx -> OK
  [5/5] converting: report5.docx -> OK
[INFO] Exiting Word instance...
[INFO] All tasks completed! Total time: 1m23s

========================================
Total: 5 | Success: 4 | Failed: 0
Duration: 1m23s
```

### 退出码

| 退出码 | 含义 |
|--------|------|
| 0 | 全部成功 |
| 1 | 运行出错（路径不存在等） |
| 2 | 部分文件转换失败 |

---

## 通用说明

### 编译

```bash
cd RTFtool_CLI
go build -o rtftool.exe .
```

### 关于 Word 进程

- 使用 Word COM 的命令（`check`、`combine-docx`、`convert`、`docx2rtf`）在运行前会自动检测并终止残留的 `WINWORD.EXE` 进程
- 如程序意外中断导致 Word 进程残留，可手动执行 `taskkill /F /IM WINWORD.EXE`

### 在脚本中使用

CLI 版本适合在 BAT/PowerShell 脚本中集成使用，例如：

```bat
@echo off
REM 步骤1：合并 RTF 文件
rtftool.exe combine -i "C:\study\rtf_source" -o "C:\study\output" -n "combined" -t -p
if %errorlevel% neq 0 (
    echo Combine failed!
    exit /b 1
)

REM 步骤2：将合并后的 RTF 转换为 PDF
rtftool.exe convert --pdf "C:\study\output\combined.rtf"
if %errorlevel% neq 0 (
    echo Convert failed!
    exit /b 1
)

echo All done!
```

```powershell
# PowerShell 示例
$source = "C:\study\rtf_source"
$output = "C:\study\output"

# 合并
& .\rtftool.exe combine -i $source -o $output -n "combined" -t -p
if ($LASTEXITCODE -ne 0) { Write-Error "Combine failed"; exit 1 }

# 转换
& .\rtftool.exe convert --pdf "$output\combined.rtf"
if ($LASTEXITCODE -ne 0) { Write-Error "Convert failed"; exit 1 }

Write-Host "All done!"
```
