# RTFtool CLI 化方案

## 一、现有 GUI 软件功能分析

RTF Tools v0.3 是一个基于 Windows `walk` GUI 库的桌面工具，通过 Microsoft Word COM (OLE) 接口操作 RTF/DOCX 文件。共包含 **5 个核心功能模块**：

### 功能 1：RTF Page Check（RTF 页码校验）

| 项目 | 说明 |
|------|------|
| **用途** | 校验指定文件夹（含子文件夹）内所有 RTF 文件的页码一致性 |
| **原理** | 通过 Word COM 获取实际渲染页数，同时从 RTF 文本中正则解析 `Page 1 of N` 标记，对比两者。不一致时自动深入到节级别定位差异 |
| **输入** | 一个文件夹路径 |
| **输出** | 每个文件的校验结果（匹配/不匹配/失败），汇总统计 |
| **特点** | 按 CPU 线程数并行处理，大文件优先调度 |

### 功能 2：RTF Combine - Specify（RTF 合并 - 指定样式）

| 项目 | 说明 |
|------|------|
| **用途** | 将多个 RTF 文件按指定顺序合并为一个，支持自动生成目录（TOC）和刷新全局页码 |
| **原理** | 纯文本层面操作 RTF 内容：提取标题、解析页码、拼接内容、生成带超链接的 TOC |
| **输入** | 多个 RTF 文件路径（有序）、是否添加 TOC、每页行数、是否刷新页码、输出路径、输出文件名 |
| **输出** | 合并后的单个 RTF 文件 |
| **特点** | 不依赖 Word COM，纯文本操作，速度快 |

### 功能 3：Docx Combine - General（Docx/RTF 通用合并）

| 项目 | 说明 |
|------|------|
| **用途** | 将多个 DOCX 或 RTF 文件按顺序合并为一个 DOCX |
| **原理** | 通过 Word COM 打开基准文档，逐一插入分节符和后续文件内容，再处理 `\pgnrestart` 控制符 |
| **输入** | 多个 DOCX/RTF 文件路径（有序）、输出路径、输出文件名 |
| **输出** | 合并后的 DOCX 文件（同时生成中间 RTF） |
| **特点** | 依赖 Word COM，不支持 TOC |

### 功能 4a：RTF → PDF / DOCX 转换

| 项目 | 说明 |
|------|------|
| **用途** | 将单个 RTF 文件转换为 PDF 和/或 DOCX |
| **原理** | 通过 Word COM 打开 RTF、修改文档属性（标题清空、作者设为 ZaiLab）、导出为 PDF/DOCX，PDF 还会调用内嵌的 `optimize_pdf.exe` 进行优化 |
| **输入** | 一个 RTF 文件路径、是否转 PDF、是否转 DOCX |
| **输出** | PDF 文件和/或 DOCX 文件 |

### 功能 4b：DOCX → RTF 转换

| 项目 | 说明 |
|------|------|
| **用途** | 将单个 DOCX 或整个文件夹的 DOCX 批量转换为 RTF |
| **原理** | 通过 Word COM 打开 DOCX 然后 SaveAs2 为 RTF 格式 |
| **输入** | 一个 DOCX 文件路径或一个文件夹路径 |
| **输出** | 对应的 RTF 文件（已存在同名 RTF 则跳过） |

### 功能 5：配置持久化

| 项目 | 说明 |
|------|------|
| **用途** | 自动保存/恢复用户上次填写的路径和选项 |
| **存储位置** | `%AppData%/RTF_Tools/config.json` |

---

## 二、CLI 化设计方案

### 2.1 技术选型

| 项目 | 方案 |
|------|------|
| **CLI 框架** | [cobra](https://github.com/spf13/cobra) — Go 生态最主流的 CLI 框架，支持子命令、flag 自动解析、自动生成帮助文档 |
| **业务逻辑复用** | 直接复用 `main.go` 中所有核心函数（`CombineRTF`、`RTFPageCheck`、`RTFConverter`、`ConvertDocxToRTF`、`CombineDocx`），这些函数与 GUI 完全解耦，接口清晰 |
| **日志输出** | CLI 模式下 `LogCallback` 直接映射为 `fmt.Printf`，输出到终端 stdout |
| **嵌入资源** | `optimize_pdf.exe` 的 `//go:embed` 逻辑保持不变 |
| **配置** | CLI 不需要持久化配置，所有参数通过命令行 flag 传入 |

### 2.2 子命令设计

```
rtftool <command> [flags]
```

#### 命令一览表

| 子命令 | 功能 | 对应 GUI Tab |
|--------|------|-------------|
| `check` | RTF 页码校验 | Tab 1 |
| `combine` | RTF 合并（Specify 样式） | Tab 2 |
| `combine-docx` | Docx/RTF 通用合并 | Tab 3 |
| `convert` | RTF → PDF/DOCX 转换 | Tab 4 上半部分 |
| `docx2rtf` | DOCX → RTF 转换 | Tab 4 下半部分 |

---

### 2.3 各子命令详细设计

#### `rtftool check` — RTF 页码校验

```
rtftool check --dir <RTF文件夹路径>
```

| Flag | 短写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--dir` | `-d` | string | 是 | 要校验的 RTF 文件夹路径（递归扫描子文件夹） |

**示例：**
```bash
rtftool check -d "C:\Projects\RTF_Output"
```

**输出示例：**
```
[INFO] Initializing environment...
[INFO] Found 15 RTF files, using 8 workers...
[Worker 0] Processing [1/15]: T-14-1-1.rtf
[Worker 1] Processing [2/15]: T-14-2-1.rtf
...
[OK]   T-14-1-1.rtf         | App: 3  | Text: 3
[WARN] T-14-2-1.rtf         | App: 5  | Text: 4  | Detail: Exception starts at section 3
[OK]   F-14-1.rtf            | App: 1  | Text: 1
...
========================================
Total: 15 | Matched: 13 | Mismatched: 1 | Failed: 1
Duration: 32.5s
```

---

#### `rtftool combine` — RTF 合并（Specify 样式）

```
rtftool combine [flags] <file1.rtf> <file2.rtf> ...
```

| Flag | 短写 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|------|--------|------|
| `--out-dir` | `-o` | string | 是 | — | 输出目录 |
| `--out-name` | `-n` | string | 是 | — | 输出文件名（不含扩展名） |
| `--toc` | `-t` | bool | 否 | false | 是否生成目录 |
| `--toc-rows` | — | int | 否 | 23 | 目录每页行数 |
| `--refresh-page` | `-p` | bool | 否 | false | 是否刷新页码 |
| `--input-dir` | `-i` | string | 否 | — | 从指定文件夹自动扫描 RTF 文件（替代手动列出文件） |
| `--sort` | `-s` | bool | 否 | true | 使用自动排序（按 T/F/L 前缀和数字排序） |

**说明：** 文件来源有两种方式，任选其一：
1. **位置参数**：直接在命令后列出所有 RTF 文件路径，按列出顺序合并
2. **`--input-dir` flag**：指定一个文件夹，自动扫描其中所有 `.rtf` 文件并按规则排序

**示例：**
```bash
# 方式1：手动指定文件和顺序
rtftool combine -o "C:\Output" -n "Combined" -t -p "C:\RTF\T-14-1.rtf" "C:\RTF\T-14-2.rtf" "C:\RTF\F-14-1.rtf"

# 方式2：扫描文件夹自动排序
rtftool combine -i "C:\RTF\source" -o "C:\Output" -n "Combined" --toc --refresh-page
```

---

#### `rtftool combine-docx` — Docx/RTF 通用合并

```
rtftool combine-docx [flags] <file1> <file2> ...
```

| Flag | 短写 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|------|--------|------|
| `--out-dir` | `-o` | string | 是 | — | 输出目录 |
| `--out-name` | `-n` | string | 是 | — | 输出文件名（不含扩展名） |
| `--input-dir` | `-i` | string | 否 | — | 从指定文件夹扫描 .docx/.rtf 文件 |
| `--sort` | `-s` | bool | 否 | true | 使用自动排序 |

**示例：**
```bash
rtftool combine-docx -o "C:\Output" -n "Merged" "C:\docs\part1.docx" "C:\docs\part2.docx"
```

---

#### `rtftool convert` — RTF → PDF/DOCX

```
rtftool convert [flags] <source.rtf>
```

| Flag | 短写 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|------|--------|------|
| `--pdf` | — | bool | 否 | false | 转换为 PDF |
| `--docx` | — | bool | 否 | false | 转换为 DOCX |

至少需要指定 `--pdf` 或 `--docx` 之一。

**示例：**
```bash
rtftool convert --pdf --docx "C:\RTF\report.rtf"
```

---

#### `rtftool docx2rtf` — DOCX → RTF

```
rtftool docx2rtf <path>
```

| 位置参数 | 说明 |
|----------|------|
| `path` | 单个 `.docx` 文件路径，或包含多个 `.docx` 的文件夹路径（批量模式） |

**示例：**
```bash
# 单文件
rtftool docx2rtf "C:\docs\report.docx"

# 批量转换
rtftool docx2rtf "C:\docs\batch_folder"
```

---

### 2.4 项目结构设计

```
RTFtool_CLI/
├── go.mod
├── go.sum
├── main.go                 # 程序入口，初始化 cobra root command
├── core.go                 # 从 RTFtool/main.go 复制的核心业务逻辑
│                           # (CombineRTF, RTFPageCheck, RTFConverter,
│                           #  ConvertDocxToRTF, CombineDocx 等所有函数)
│                           # 移除 GUI 相关的 import，LogCallback 保留
├── cmd/
│   ├── root.go             # cobra root 命令定义
│   ├── check.go            # check 子命令
│   ├── combine.go          # combine 子命令
│   ├── combine_docx.go     # combine-docx 子命令
│   ├── convert.go          # convert 子命令
│   └── docx2rtf.go         # docx2rtf 子命令
├── optimize_pdf.exe        # 嵌入的 PDF 优化器（同原项目）
└── README.md               # CLI 使用说明
```

### 2.5 核心代码复用策略

GUI 版本的 `main.go` 中所有核心函数已经与 GUI 解耦，可以直接复用：

| 核心函数 | 改动 |
|----------|------|
| `CombineRTF()` | **无需改动**，参数全部通过函数入参传入 |
| `RTFPageCheck()` | **无需改动**，LogCallback 在 CLI 中传入 `fmt.Printf` 即可 |
| `RTFConverter()` | **无需改动** |
| `ConvertDocxToRTF()` | **无需改动** |
| `CombineDocx()` | **无需改动** |
| `KillWordProcesses()` | **无需改动** |
| `OptimizePDFWithExe()` | **无需改动** |
| 各种辅助函数 | **无需改动** |

唯一需要删除的是：
- `gui.go` 整个文件（所有 `walk` 相关代码）
- `main()` 中的 `runSimpleGUI()` 调用，替换为 cobra 的 `Execute()`
- `go.mod` 中移除 `github.com/lxn/walk` 和 `github.com/lxn/win` 依赖
- 新增 `github.com/spf13/cobra` 依赖

### 2.6 新增的 CLI 特有功能

在 GUI 中有一些交互式操作（如文件选择、排序），在 CLI 中需要用其他方式替代：

| GUI 操作 | CLI 替代方案 |
|----------|-------------|
| 文件夹浏览对话框 | 通过 `--dir` / `--input-dir` flag 直接传入路径 |
| 文件勾选与排序 | 通过位置参数按列出顺序排列，或通过 `--input-dir` + `--sort` 自动排序 |
| 进度条 | 通过终端日志文字输出进度 |
| 弹窗提示 | 通过 stdout 输出结果并使用退出码表示成功/失败 |
| 配置持久化 | 不需要，所有参数由命令行传入 |

### 2.7 退出码约定

| 退出码 | 含义 |
|--------|------|
| 0 | 成功 |
| 1 | 一般错误（参数错误、文件不存在等） |
| 2 | 部分文件处理失败（如 check 命令有不匹配的文件） |

---

## 三、实施步骤

1. **创建项目骨架**：在 `RTFtool_CLI/` 下初始化 Go module、创建目录结构
2. **复制核心逻辑**：将 `RTFtool/main.go` 中除 `main()` 函数外的所有代码复制到 `core.go`
3. **实现 cobra 命令**：逐个实现 5 个子命令
4. **实现文件扫描与排序**：将 GUI 中 `showSelectedFiles` 的排序逻辑提取为独立函数，供 `--input-dir` + `--sort` 使用
5. **编译测试**：`go build -o rtftool.exe .`
6. **编写 README**

---

## 四、依赖对比

| 依赖 | GUI 版 | CLI 版 |
|------|--------|--------|
| `github.com/go-ole/go-ole` | 需要 | 需要（Word COM 操作） |
| `github.com/lxn/walk` | 需要 | **不需要** |
| `github.com/lxn/win` | 需要 | **不需要** |
| `github.com/spf13/cobra` | 不需要 | **新增** |
| `golang.org/x/sys` | 间接 | 间接（go-ole 需要） |

---

## 五、补充说明

- CLI 版本仅支持 **Windows**（依赖 Word COM 和 `taskkill` 等 Windows 特有 API）
- `optimize_pdf.exe` 通过 `//go:embed` 嵌入，与 GUI 版行为一致
- 所有 Word COM 功能（check、combine-docx、convert、docx2rtf）执行前会自动 kill 残留的 WINWORD.EXE 进程
- `combine`（Specify 样式）不依赖 Word COM，纯文本操作
