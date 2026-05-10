package main

import (
	"archive/zip"
	"bytes"
	_ "embed"
	"fmt"
	"io"
	"log"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"syscall"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// ==============================================================================
// 1. 类型与全局常量定义 (Types & Constants)
// ==============================================================================

// LogCallback 日志回调函数类型，用于统一处理日志输出
type LogCallback func(format string, args ...interface{})

// Word COM 常量定义 (避免代码中的魔法数字)
const (
	wdCollapseEnd            = 0
	wdSectionBreakNextPage   = 2
	wdActiveEndPageNumber    = 3
	wdFormatRTF              = 6
	wdFormatDocumentDefault  = 16
	wdExportFormatPDF        = 17
	wdExportOptimizeForPrint = 0
	wdExportAllDocument      = 0
	wdExportDocumentContent  = 0
)

// CombineFileInfo 用于 RTF 合并的文件信息结构
type CombineFileInfo struct {
	Name     string
	Path     string
	Ord      int
	Filename string
	Name1    string
	Order    int
	Title    string
	Page     int
	PageNum  int
	BodyOrd  int
	Title1   string
	SortNums []int
	Content  string
}

// RTFPageCheckFileInfo 用于 RTF 页码检查的文件信息
type RTFPageCheckFileInfo struct {
	path string
	size int64
}

// job 内部任务结构
type job struct {
	filePath string
	index    int
}

// rtfPageCheckResult 内部单文件检查结果
type rtfPageCheckResult struct {
	filePath       string
	pageCountApp   int
	pageCountText  int
	mismatchDetail string
	err            error
	index          int
}

// RTFPageCheckResult 最终整体检查结果
type RTFPageCheckResult struct {
	TotalFiles              int
	SuccessCount            int
	FailedCount             int
	AllMatched              bool
	Duration                time.Duration
	RTFPageCheckFileResults []RTFPageCheckFileResult
	Error                   string
}

// RTFPageCheckFileResult 导出的单文件检查结果
type RTFPageCheckFileResult struct {
	FilePath       string
	PageCountApp   int
	PageCountText  int
	Matched        bool
	MismatchDetail string
	Error          string
}

// ConversionResult 格式转换结果
type ConversionResult struct {
	TotalFiles   int
	SuccessCount int
	ErrorCount   int
	Duration     time.Duration
	Error        string
}

type parsedWordDocument struct {
	Prefix      string
	BodyContent string
	Suffix      string
}

//go:embed optimize_pdf.exe
var pdfOptimizerExe []byte

var (
	bodyPattern            = regexp.MustCompile(`(?s)\A(.*?<w:body[^>]*>)(.*?)(</w:body>.*)</w:document>\s*\z`)
	sectPrSelfPattern      = regexp.MustCompile(`(?s)\A<w:sectPr\b[^>]*/>`)
	sectPrFullPattern      = regexp.MustCompile(`(?s)\A<w:sectPr\b.*?</w:sectPr>`)
	headerFooterRefPattern = regexp.MustCompile(`(?s)<w:(?:headerReference|footerReference)\b[^>]*/>`)
	settingsSelfPattern    = regexp.MustCompile(`(?s)<w:settings\b([^>]*)/>`)
	updateFieldsPattern    = regexp.MustCompile(`(?s)<w:updateFields\b[^>]*/>|<w:updateFields\b.*?</w:updateFields>`)
	relationshipTagPattern = regexp.MustCompile(`(?s)<Relationship\b[^>]*/>`)
	overrideTagPattern     = regexp.MustCompile(`(?s)<Override\b[^>]*/>`)
	rIDPattern             = regexp.MustCompile(`\brId(\d+)\b`)
)

const defaultTOCTitle = "Table of Contents"
const defaultTOCFieldCode = `TOC \o "1-3" \h \z \u`

const (
	settingsPath             = "word/settings.xml"
	documentRelsPath         = "word/_rels/document.xml.rels"
	contentTypesPath         = "[Content_Types].xml"
	settingsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	settingsContentType      = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	updateFieldsElement      = `<w:updateFields w:val="true"/>`
	defaultSettingsXML       = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + updateFieldsElement + `</w:settings>`
	defaultRelationshipsXML  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
	defaultSettingsOverride  = `<Override PartName="/word/settings.xml" ContentType="` + settingsContentType + `"/>`
)

// ==============================================================================
// 2. 通用辅助函数 (Shared Utility Functions)
// ==============================================================================

// KillWordProcesses 统一终止后台僵尸 Word 进程
func KillWordProcesses() error {
	fmt.Println("🔍 Checking for WINWORD.EXE process....")

	checkCmd := exec.Command("tasklist", "/FI", "IMAGENAME eq WINWORD.EXE")
	checkCmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}

	output, _ := checkCmd.CombinedOutput()
	outputStr := strings.ToLower(string(output))

	if !strings.Contains(outputStr, "winword.exe") {
		fmt.Println("✅ No WINWORD.EXE process found. No need to terminate.")
		return nil
	}

	fmt.Println("⚠️ WINWORD.EXE process detected. Attempting to terminate....")

	killCmd := exec.Command("taskkill", "/F", "/IM", "WINWORD.EXE")
	killCmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}

	killOutput, killErr := killCmd.CombinedOutput()
	if killErr != nil {
		return fmt.Errorf("Taskkill failed:  %v, Output: %s", killErr, string(killOutput))
	}

	fmt.Println("✅ WINWORD.EXE Process terminated successfully.")
	return nil
}

// ==============================================================================
// 3. RTF 合并模块 (Combine RTF Module)
// ==============================================================================

func processRTFContent(inputFile string) (string, error) {
	data, err := os.ReadFile(inputFile) // 替换 ioutil
	if err != nil {
		return "", fmt.Errorf("Read file failed %s: %w", inputFile, err)
	}

	originalContent := string(data)
	lines := strings.Split(originalContent, "\n")
	var processedLines []string

	for _, line := range lines {
		if matched, _ := regexp.MatchString(`IDX\d+`, line); matched {
			fmt.Printf("Found IDX line: %s\n", line)
			// 保留原有的被注释掉的逻辑供未来参考
		}
		processedLines = append(processedLines, line)
	}

	return strings.Join(processedLines, "\n"), nil
}

func rtfChineseEncoder(chineseText string) string {
	var rtfBody strings.Builder
	for _, char := range chineseText {
		if char > 127 {
			rtfBody.WriteString(fmt.Sprintf("\\u%d;", char))
		} else {
			rtfBody.WriteRune(char)
		}
	}
	return fmt.Sprintf("{\\cf0\\b %s \\b0\\tab}", rtfBody.String())
}

func grabFirstDocPageSize(content string) (int, int) {
	const defaultWidth = 15840
	const defaultHeight = 12240

	if content == "" {
		return defaultWidth, defaultHeight
	}

	lines := strings.Split(content, "\n")
	for _, line := range lines {
		if strings.Contains(line, "pgwsxn") {
			re := regexp.MustCompile(`pgwsxn(\d+)[^\\]*\\pghsxn(\d+)`)
			match := re.FindStringSubmatch(line)
			if len(match) >= 3 {
				width, err1 := strconv.Atoi(match[1])
				height, err2 := strconv.Atoi(match[2])
				if err1 == nil && err2 == nil {
					return width, height
				}
			}
			break
		}
	}

	return defaultWidth, defaultHeight
}

// CombineRTF Combine multiple RTF files to one.
func CombineRTF(srcPath []string, addtoc bool, rowOfTocInAPage int, changePage bool, outPath, outFile string) error {
	startTime := time.Now()

	// 1. 获取文件列表并预处理内容
	var files []CombineFileInfo
	for _, filePath := range srcPath {
		if !strings.HasSuffix(strings.ToLower(filePath), ".rtf") {
			continue
		}

		fileInfo, err := os.Stat(filePath)
		if err != nil {
			return fmt.Errorf("Failed to read file attributes: %v", err)
		}
		if fileInfo.IsDir() {
			continue
		}

		processedContent, err := processRTFContent(filePath)
		if err != nil {
			log.Printf("Skipping failed file: %v", err)
			continue
		}

		files = append(files, CombineFileInfo{
			Name:    filepath.Base(filePath),
			Path:    filepath.Dir(filePath),
			Content: processedContent,
		})
	}

	for i := range files {
		files[i].Order = i + 1
	}

	pageAdd := int(math.Ceil(float64(len(files)) / float64(rowOfTocInAPage)))
	fmt.Println("Page additions for TOC:", pageAdd)

	// 2. 获取页面尺寸
	pgwsxnStyle := []int{15840, 12240}
	if len(files) > 0 && files[0].Content != "" {
		pgwsxnStyle[0], pgwsxnStyle[1] = grabFirstDocPageSize(files[0].Content)
	}
	fmt.Printf("Page Size: %v\n", pgwsxnStyle)

	// 3. 提取标题生成目录
	var titles []string
	for _, file := range files {
		content := ""
		if file.Content == "" {
			titles = append(titles, "")
			continue
		}

		lines := strings.Split(file.Content, "\n")
		var prevLine, idxLine, nextLine string
		for i, line := range lines {
			if strings.Contains(line, "IDX") {
				idxLine = strings.TrimSpace(line)
				if i > 0 {
					prevLine = strings.TrimSpace(lines[i-1])
				}
				if i < len(lines)-1 {
					nextLine = strings.TrimSpace(lines[i+1])
				}
				break
			}
		}

		if idxLine != "" {
			content = fmt.Sprintf("Previous Line: %s\nIDX Line: %s\nNext Line: %s\n\n", prevLine, idxLine, nextLine)
			content = strings.Split(content, `\line{`)[0]
			if parts := strings.Split(content, `\s999 \b `); len(parts) > 1 {
				content = parts[1]
			}
			content = strings.Split(content, `\b`)[0]

			content = strings.ReplaceAll(content, `\line`, " ")
			content = regexp.MustCompile(`\s+`).ReplaceAllString(content, " ")
			content = strings.TrimSpace(content)
		} else {
			log.Printf("IDX not found, ERROR in %s\n", file.Name)
		}
		titles = append(titles, content)
	}

	for i := range titles {
		title := strings.TrimSpace(titles[i])
		title = strings.ReplaceAll(title, "\\s999", "")
		title = strings.ReplaceAll(title, "\\b0", "")
		title = strings.ReplaceAll(title, "\\b", "")
		title = strings.ReplaceAll(title, "{\\par}", "")
		titles[i] = fmt.Sprintf("{\\cf0\\b %s \\b0\\tab}", strings.TrimSpace(title))
	}

	// 4. 提取与处理页码 (升级版)
	var pages []int
	for _, file := range files {
		if file.Content == "" {
			pages = append(pages, 0)
			continue
		}

		pageMatched := false
		lines := strings.Split(file.Content, "\n")

		reNoise := regexp.MustCompile(`[\\{}]|(\\[a-zA-Z]+\s?)`)
		reExtractNum := regexp.MustCompile(`(?i)Page\s*1\s*(?:of|/)\s*(\d+)`)

		for _, line := range lines {
			if strings.Contains(line, "Page") {
				cleanLine := reNoise.ReplaceAllString(line, "")
				cleanLine = strings.TrimSpace(cleanLine)
				match := reExtractNum.FindStringSubmatch(cleanLine)
				if len(match) == 2 {
					if num, err := strconv.Atoi(match[1]); err == nil {
						pages = append(pages, num)
						pageMatched = true
						break
					}
				}
			}
		}
		if !pageMatched {
			pages = append(pages, 0)
		}
	}

	var pageNum []int
	cumSum := 0
	for i, page := range pages {
		pageNum = append(pageNum, cumSum+1+pageAdd)
		cumSum += page
		files[i].Title = titles[i]
		files[i].Page = page
		files[i].PageNum = pageNum[i]
	}

	// 5. 生成主干内容
	var contentBuilder strings.Builder
	for i, file := range files {
		text := file.Content
		if text == "" {
			continue
		}

		for _, splitPoint := range []string{"\\widowctrl", "\\sectd", "\\info"} {
			if strings.Contains(text, splitPoint) {
				text = splitPoint + strings.SplitN(text, splitPoint, 2)[1]
				break
			}
		}

		text = strings.TrimSpace(text)
		text = strings.ReplaceAll(text, "IDX", fmt.Sprintf("IDX%d", file.Order))
		text = strings.TrimSuffix(text, "}")

		contentBuilder.WriteString(text)
		if i != len(files)-1 {
			contentBuilder.WriteString("\\sect")
		}
	}

	// 6. 生成头部与目录
	var header string
	if len(files) > 0 && files[0].Content != "" {
		text := files[0].Content
		for _, splitPoint := range []string{"\\widowctrl", "\\sectd", "\\info"} {
			if strings.Contains(text, splitPoint) {
				header = strings.Split(text, splitPoint)[0]
				break
			}
		}
	}

	toc := fmt.Sprintf("\\sectd\\linex0\\endnhere\\pgwsxn%d\\pghsxn%d\\lndscpsxn\\pgnrestart"+
		"\\pgnstarts1\\headery1440\\footery1440\\marglsxn1440\\margrsxn1440\\margtsxn1440"+
		"\\margbsxn1440\n\\qc\\f36\\b\\s0{\\s999 Table of Contents \\par}\\s0\\b0\\par\\par\\pard\n",
		pgwsxnStyle[0], pgwsxnStyle[1])

	currentPageRows := 0
	for _, file := range files {
		titleRTF := rtfChineseEncoder(file.Title)
		toc += fmt.Sprintf("\\fi-2000\\li2000{\\f1\\fs18\\cf0{\\field{\\*\\fldinst { HYPERLINK \\\\l \"IDX%d\"}}{\\fldrslt%s}}\n\\ptabldot\\pindtabqr{\\field{ %d }}}\\s0\\par\n",
			file.Order, titleRTF, file.PageNum)
		currentPageRows++
		if currentPageRows >= rowOfTocInAPage {
			toc += "\\pard\\sect"
			currentPageRows = 0
		}
	}
	toc += "\\pard\\sect"

	// 7. 拼接最终内容
	var finalContent string
	if addtoc {
		finalContent = header + toc + contentBuilder.String() + "}"
	} else {
		finalContent = header + contentBuilder.String() + "}"
	}

	if !strings.HasPrefix(finalContent, "{\\rtf1") {
		finalContent = "{\\rtf1\\ansi\\ansicpg936\\deff0\\deflang1033\\deflangfe2052{\\fonttbl{\\f0\\fnil\\fcharset134 \\'cb\\'ce\\'cc\\'e5;}}\n" +
			"{\\*\\generator Msftedit 5.41.21.2510;}\n\\viewkind4\\uc1\\pard\\lang2052\\f0\\fs18\n" + finalContent
	}
	if !strings.HasSuffix(finalContent, "}") {
		finalContent += "}"
	}

	// 8. 页码处理逻辑 (升级版)
	if len(files) > 0 {
		lastFile := files[len(files)-1]
		totalPage := lastFile.Page + lastFile.PageNum - 1

		if changePage {
			// 核心正则拆解为 4 个捕获组：
			// 组1: "Page " 及可能跟随的格式符
			// 组2: 原始当前页码 (X)
			// 组3: 中间的连接符，包含 of, /, 以及前后的空格、括号等 RTF 噪音
			// 组4: 原始总页码 (Y)
			re := regexp.MustCompile(`(?i)(Page\s*[\\{\}]*\s*)(\d+)(\s*[\\{\}]*\s*(?:of|/)\s*[\\{\}]*\s*)(\d+)`)

			counter := 1 + pageAdd
			finalContent = re.ReplaceAllStringFunc(finalContent, func(match string) string {
				// 获取所有捕获组的内容
				submatches := re.FindStringSubmatch(match)

				// submatches[1] 是前缀 (例如 "Page {")
				// submatches[3] 是中间的连接符 (例如 "} of {" 或 " / ")
				// 我们保留原有的 RTF 格式结构，只把中间的组2和组4替换为新的数字
				result := fmt.Sprintf("%s%d%s%d", submatches[1], counter, submatches[3], totalPage)

				counter++
				return result
			})
		}
		rePgn := regexp.MustCompile(`\\pgnrestart\\pgnstarts\d+`)
		indices := rePgn.FindAllStringIndex(finalContent, -1)
		if len(indices) > 1 {
			for i := len(indices) - 1; i >= 1; i-- {
				finalContent = finalContent[:indices[i][0]] + finalContent[indices[i][1]:]
			}
		}

		if err := os.MkdirAll(outPath, 0755); err != nil {
			return fmt.Errorf("Failed to create output directory: %v", err)
		}

		outputPath := filepath.Join(outPath, outFile+".rtf")
		fmt.Printf("Writing file: %s, Length: %d\n", outputPath, len(finalContent))

		if err := os.WriteFile(outputPath, []byte(finalContent), 0644); err != nil {
			return fmt.Errorf("Failed to writing file: %v", err)
		}
		fmt.Printf("Write file successfully: %s\n", outputPath)
	}

	fmt.Printf("Combine finished! Time Taken: %.1fs\n", time.Since(startTime).Seconds())
	return nil
}

// ==============================================================================
// 4. RTF 页码检查模块 (RTF Page Check)
// ==============================================================================

func RTFPageCheck(rtfDir string, logCallback LogCallback) *RTFPageCheckResult {
	start := time.Now()
	rtfDir = strings.TrimSpace(rtfDir)

	logCallback("🔍 Initializing environment......\n")
	_ = KillWordProcesses()

	fileInfos, err := findRtfFiles(rtfDir)
	if err != nil {
		logCallback("❌ Failed to find file: %v\n", err)
		return &RTFPageCheckResult{Error: err.Error()}
	}

	totalFiles := len(fileInfos)
	if totalFiles == 0 {
		logCallback("⚠️ No valid .rtf files found. \n")
		return &RTFPageCheckResult{TotalFiles: 0}
	}

	sort.Slice(fileInfos, func(i, j int) bool { return fileInfos[i].size > fileInfos[j].size })

	numWorkers := runtime.NumCPU()
	if numWorkers > totalFiles {
		numWorkers = totalFiles
	}

	jobs := make(chan job, totalFiles)
	results := make(chan rtfPageCheckResult, totalFiles)
	var wg sync.WaitGroup

	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go rtfCheckWorker(i, jobs, results, &wg, totalFiles, logCallback)
	}

	for i, info := range fileInfos {
		jobs <- job{filePath: info.path, index: i}
	}
	close(jobs)
	wg.Wait()
	close(results)

	return collectCheckResults(results, totalFiles, start, logCallback)
}

func rtfCheckWorker(workerID int, jobs <-chan job, results chan<- rtfPageCheckResult, wg *sync.WaitGroup, totalFiles int, logCallback LogCallback) {
	defer wg.Done()
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		logCallback("❌ Thread %d failed to start Word: %v\n", workerID, err)
		return
	}
	wordApp, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer wordApp.Release()

	oleutil.PutProperty(wordApp, "Visible", false)
	oleutil.PutProperty(wordApp, "DisplayAlerts", 0)

	for j := range jobs {
		logCallback("👷 [Worker %d] Processing  [ %d/%d ] : %s\n", workerID, j.index+1, totalFiles, filepath.Base(j.filePath))
		results <- processSingleFile(wordApp, j.filePath, j.index, logCallback)
	}
	oleutil.CallMethod(wordApp, "Quit")
}

func processSingleFile(wordApp *ole.IDispatch, filePath string, index int, log LogCallback) rtfPageCheckResult {
	res := rtfPageCheckResult{filePath: filePath, index: index}

	textCount, err := getPageCountFromRtfText(filePath)
	if err != nil {
		res.err = err
		return res
	}
	res.pageCountText = textCount

	absPath, _ := filepath.Abs(filePath)
	documents := oleutil.MustGetProperty(wordApp, "Documents").ToIDispatch()
	defer documents.Release()

	docVariant, err := oleutil.CallMethod(documents, "Open", absPath)
	if err != nil {
		res.err = fmt.Errorf("Word can not open this file.")
		return res
	}
	doc := docVariant.ToIDispatch()
	defer doc.Release()

	selection := oleutil.MustGetProperty(wordApp, "Selection").ToIDispatch()
	defer selection.Release()

	res.pageCountApp = int(oleutil.MustGetProperty(selection, "Information", 4).Val)

	if res.pageCountApp != res.pageCountText {
		log("  ⚠️ Page number mismatch! [App: %d, Text: %d] Deep scanning...\n", res.pageCountApp, res.pageCountText)
		res.mismatchDetail = performDeepAlignmentCheck(wordApp, doc)
	}

	oleutil.CallMethod(doc, "Close", 0)
	return res
}

func performDeepAlignmentCheck(wordApp, doc *ole.IDispatch) string {
	sections := oleutil.MustGetProperty(doc, "Sections").ToIDispatch()
	defer sections.Release()
	sectionCount := int(oleutil.MustGetProperty(sections, "Count").Val)

	selection := oleutil.MustGetProperty(wordApp, "Selection").ToIDispatch()
	defer selection.Release()

	for i := 1; i <= sectionCount; i++ {
		secItem, err := oleutil.CallMethod(sections, "Item", i)
		if err != nil {
			continue
		}
		sec := secItem.ToIDispatch()
		secRange := oleutil.MustGetProperty(sec, "Range").ToIDispatch()
		startPos := oleutil.MustGetProperty(secRange, "Start").Val

		oleutil.CallMethod(selection, "SetRange", startPos, startPos)
		pageVal, _ := oleutil.GetProperty(selection, "Information", wdActiveEndPageNumber)
		actualPage := int(pageVal.Val)

		sec.Release()
		secRange.Release()

		if i != actualPage {
			return fmt.Sprintf("Exception starts at section %d (Page %d )", i, actualPage-1)
		}
	}
	return "No section start offset found. Difference may be caused by trailing blank pages."
}

func getPageCountFromRtfText(filePath string) (int, error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return 0, fmt.Errorf("Read filed: %v", err)
	}
	re := regexp.MustCompile(`(?i)Page\s*1.*?(\d+)`)
	matches := re.FindStringSubmatch(string(content))
	if len(matches) < 2 {
		return 0, fmt.Errorf("No recognizable page number markers(Page X of X) found.")
	}
	return strconv.Atoi(matches[1])
}

func findRtfFiles(dir string) ([]RTFPageCheckFileInfo, error) {
	var files []RTFPageCheckFileInfo
	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err == nil && !info.IsDir() && strings.HasSuffix(strings.ToLower(info.Name()), ".rtf") {
			if !strings.HasPrefix(info.Name(), "~") {
				files = append(files, RTFPageCheckFileInfo{path: path, size: info.Size()})
			}
		}
		return nil
	})
	return files, err
}

func collectCheckResults(results chan rtfPageCheckResult, total int, start time.Time, log LogCallback) *RTFPageCheckResult {
	final := &RTFPageCheckResult{TotalFiles: total, AllMatched: true}
	var resList []rtfPageCheckResult
	for i := 0; i < total; i++ {
		resList = append(resList, <-results)
	}
	sort.Slice(resList, func(i, j int) bool { return resList[i].index < resList[j].index })

	for _, r := range resList {
		matched := r.err == nil && r.pageCountApp == r.pageCountText

		if r.err != nil {
			final.FailedCount++
			final.AllMatched = false
			log("❌ Failed: %-25s | Error: %v\n", filepath.Base(r.filePath), r.err)
		} else if !matched {
			final.AllMatched = false
			log("🚨 Error: %-25s | App: %d | Text: %d | Detail: %s\n", filepath.Base(r.filePath), r.pageCountApp, r.pageCountText, r.mismatchDetail)
		} else {
			final.SuccessCount++
		}

		final.RTFPageCheckFileResults = append(final.RTFPageCheckFileResults, RTFPageCheckFileResult{
			FilePath:       r.filePath,
			PageCountApp:   r.pageCountApp,
			PageCountText:  r.pageCountText,
			Matched:        matched,
			MismatchDetail: r.mismatchDetail,
			Error:          fmt.Sprint(r.err),
		})
	}
	final.Duration = time.Since(start)
	return final
}

// ==============================================================================
// 5. RTF Converter (PDF/DOCX)
// ==============================================================================

func OptimizePDFWithExe(inputPath string, logCallback LogCallback) error {
	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		return fmt.Errorf("Input file not found: %s", inputPath)
	}

	logCallback("🗜️ Optimizing PDF...\n")

	exeDir, err := os.Executable()
	if err != nil {
		return err
	}
	exeDir = filepath.Dir(exeDir)
	optimizerPath := filepath.Join(exeDir, "optimize_pdf.exe")

	if _, err := os.Stat(optimizerPath); os.IsNotExist(err) {
		if err := os.WriteFile(optimizerPath, pdfOptimizerExe, 0755); err != nil {
			return fmt.Errorf("Failed to extract optimizer: %v", err)
		}
		time.Sleep(100 * time.Millisecond)
	}

	if err := validateExecutable(optimizerPath); err != nil {
		return err
	}

	outputPath := strings.Replace(inputPath, "_.pdf", ".pdf", 1)
	cmd := exec.Command(optimizerPath, inputPath, outputPath)
	cmd.Dir = exeDir

	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	output, err := cmd.Output()

	if err != nil || !strings.Contains(string(output), "SUCCESS") {
		return fmt.Errorf("Optimization failed: %v, Output: %s", err, stderr.String())
	}

	_ = os.Remove(inputPath)
	logCallback("✅ PDF optimized.\n")
	return nil
}

func validateExecutable(path string) error {
	file, err := os.Open(path)
	if err != nil {
		return err
	}
	defer file.Close()

	header := make([]byte, 2)
	if _, err := file.Read(header); err != nil || header[0] != 'M' || header[1] != 'Z' {
		return fmt.Errorf("Invalid executable signature.")
	}
	return nil
}

func RTFConverter(originalRtf string, Trans_pdf bool, Trans_docx bool, logCallback LogCallback) error {
	start := time.Now()
	logCallback("🚀 Starting conversion: PDF=%t, DOCX=%t\n", Trans_pdf, Trans_docx)

	_ = KillWordProcesses()
	time.Sleep(500 * time.Millisecond)

	// 1. 拷贝临时文件
	base := filepath.Base(originalRtf)
	copyPath := filepath.Join(filepath.Dir(originalRtf), "Copy_"+base)
	logCallback("📄 Preparing temporary file: %s\n", filepath.Base(copyPath))

	srcData, err := os.ReadFile(originalRtf)
	if err != nil {
		logCallback("❌ Failed to read source file.\n")
		return fmt.Errorf("Failed to read source file: %w", err)
	}
	if err := os.WriteFile(copyPath, srcData, 0644); err != nil {
		logCallback("❌ Failed to create temporary file.\n")
		return fmt.Errorf("Failed to create temporary file: %w", err)
	}

	// 2. 执行转换
	logCallback("⚙️ Running Word conversion...\n")
	pdfPath, err := modifyAndConvertDoc(copyPath, Trans_pdf, Trans_docx, logCallback)
	if err != nil {
		logCallback("❌ Conversion failed in Word automation.\n")
		return err
	}

	_ = KillWordProcesses()

	// 3. 清理
	_ = os.Remove(copyPath)

	if Trans_pdf {
		logCallback("✅ Final PDF saved at:: %s\n", strings.Replace(pdfPath, "_.pdf", ".pdf", 1))
	}
	logCallback("Total time: %s\n", time.Since(start))
	return nil
}

// 提取出的独立 OLE 转换逻辑
func modifyAndConvertDoc(copyRtfPath string, transPdf, transDocx bool, logCallback LogCallback) (string, error) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	if err != nil && !strings.Contains(err.Error(), "already") {
		_ = ole.CoInitialize(0)
	}
	defer ole.CoUninitialize()

	word, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		return "", fmt.Errorf("Failed to create Word object: %w", err)
	}
	defer word.Release()

	oleWord := word.MustQueryInterface(ole.IID_IDispatch)
	defer oleWord.Release()
	oleutil.PutProperty(oleWord, "Visible", false)

	docs := oleutil.MustGetProperty(oleWord, "Documents").ToIDispatch()
	defer docs.Release()
	logCallback("📖 Opening temporary document in Word...\n")
	doc, err := oleutil.CallMethod(docs, "Open", copyRtfPath)
	if err != nil {
		return "", err
	}
	docDisp := doc.ToIDispatch()
	defer docDisp.Release()

	// 修改属性
	props := oleutil.MustGetProperty(docDisp, "BuiltInDocumentProperties").ToIDispatch()
	defer props.Release()

	if titleProp, err := oleutil.GetProperty(props, "Item", 1); err == nil {
		oleutil.PutProperty(titleProp.ToIDispatch(), "Value", "")
	}
	if authorProp, err := oleutil.GetProperty(props, "Item", 3); err == nil {
		oleutil.PutProperty(authorProp.ToIDispatch(), "Value", "ZaiLab")
	}
	oleutil.CallMethod(docDisp, "Save")

	// 导出逻辑
	dir, base := filepath.Dir(copyRtfPath), filepath.Base(copyRtfPath)
	cleanBase := strings.TrimSuffix(strings.TrimPrefix(base, "Copy_"), filepath.Ext(base))
	var pdfPath string

	if transPdf {
		pdfPath = filepath.Join(dir, cleanBase+"_.pdf")
		logCallback("📝 Exporting PDF...\n")
		_, err = oleutil.CallMethod(docDisp, "ExportAsFixedFormat",
			pdfPath, wdExportFormatPDF, 0, wdExportOptimizeForPrint,
			wdExportAllDocument, 0, 0, wdExportDocumentContent,
			true, true, 1, true, false, true,
		)
		if err == nil {
			time.Sleep(1 * time.Second)
			if optErr := OptimizePDFWithExe(pdfPath, logCallback); optErr != nil {
				logCallback("⚠️ PDF optimization skipped: %v\n", optErr)
			}
		} else {
			logCallback("❌ PDF export failed.\n")
		}
	}

	if transDocx {
		docxPath := filepath.Join(dir, cleanBase+".docx")
		logCallback("📝 Exporting DOCX...\n")
		_, docxErr := oleutil.CallMethod(docDisp, "SaveAs", docxPath, wdFormatDocumentDefault)
		if docxErr != nil {
			logCallback("❌ DOCX export failed: %v\n", docxErr)
		} else {
			logCallback("✅ DOCX saved: %s\n", docxPath)
		}
	}

	oleutil.CallMethod(docDisp, "Close", false)
	return pdfPath, nil
}

// ==============================================================================
// 6. Docx to RTF Module
// ==============================================================================

func ConvertDocxToRTF(inputPath string, logCallback LogCallback) ConversionResult {
	result := ConversionResult{}
	start := time.Now()

	logCallback("🔍 Scanning target path....\n")

	var docxFiles []string
	info, err := os.Stat(inputPath)
	if err != nil {
		logCallback("❌ Failed to get path info.: %v\n", err)
		return ConversionResult{Error: err.Error()}
	}

	if info.IsDir() {
		files, _ := os.ReadDir(inputPath)
		for _, f := range files {
			if !f.IsDir() && strings.HasSuffix(strings.ToLower(f.Name()), ".docx") && !strings.HasPrefix(f.Name(), "~$") {
				docxFiles = append(docxFiles, filepath.Join(inputPath, f.Name()))
			}
		}
	} else if strings.HasSuffix(strings.ToLower(inputPath), ".docx") {
		docxFiles = append(docxFiles, inputPath)
	}

	result.TotalFiles = len(docxFiles)
	if result.TotalFiles == 0 {
		logCallback("⚠️ No valid .docx files found at the specified path. \n")
		return ConversionResult{Error: "no files found"}
	}

	logCallback("✅ Found %d file(s) to convert.\n", result.TotalFiles)

	logCallback("🔌 Cleaning up background Word processes...\n")
	_ = KillWordProcesses()
	time.Sleep(500 * time.Millisecond)

	logCallback("⚙️ Initializing Word COM components....\n")
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, _ := oleutil.CreateObject("Word.Application")
	word, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer word.Release()

	oleutil.PutProperty(word, "Visible", false)
	oleutil.PutProperty(word, "DisplayAlerts", 0)

	docs := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
	defer docs.Release()

	// 循环处理每个文件
	for i, docxPath := range docxFiles {
		fileName := filepath.Base(docxPath)
		rtfPath := strings.TrimSuffix(docxPath, filepath.Ext(docxPath)) + ".rtf"

		logCallback("  [%d/%d] convert to: %s -> ", i+1, result.TotalFiles, fileName)

		if _, err := os.Stat(rtfPath); err == nil {
			logCallback("Skipped (identical RTF already exists).\n")
			continue
		}

		if docObj, err := oleutil.CallMethod(docs, "Open", docxPath); err == nil {
			doc := docObj.ToIDispatch()
			oleutil.CallMethod(doc, "SaveAs2", rtfPath, wdFormatRTF)
			oleutil.CallMethod(doc, "Close", 0)
			doc.Release()
			result.SuccessCount++
			logCallback("Successfully! ✅\n")
		} else {
			result.ErrorCount++
			logCallback("Failed ❌ (%v)\n", err)
		}
	}

	logCallback("🧹 Exiting Word instance...\n")
	oleutil.CallMethod(word, "Quit")

	result.Duration = time.Since(start)
	logCallback("🎉 All tasks completed! Total time taken: %v\n", result.Duration)
	return result
}

// ==============================================================================
// 7. Combine Docx Module
// ==============================================================================

func AddTOCToDocx(sourcePath, outputPath string) error {
	return addTOCByOpenXML(sourcePath, outputPath)
}

func addTOCByOpenXML(sourcePath, outputPath string) error {
	sourcePath = strings.TrimSpace(sourcePath)
	outputPath = strings.TrimSpace(outputPath)
	if sourcePath == "" || outputPath == "" {
		return fmt.Errorf("source and output path are required")
	}
	if strings.ToLower(filepath.Ext(sourcePath)) != ".docx" || strings.ToLower(filepath.Ext(outputPath)) != ".docx" {
		return fmt.Errorf("source/output must be .docx")
	}

	entries, err := readDocxEntries(sourcePath)
	if err != nil {
		return err
	}
	doc, err := parseWordDocument(entries["word/document.xml"])
	if err != nil {
		return fmt.Errorf("parse source document: %w", err)
	}

	tocSectionBreak := buildTOCSectionBreakParagraphWithoutHeaderFooter(entries["word/document.xml"])
	parts := []string{
		buildTextParagraph(defaultTOCTitle),
		buildTOCFieldParagraph(`TOC \o "1-3" \h \z \u`),
	}
	if tocSectionBreak != "" {
		parts = append(parts, tocSectionBreak)
	} else {
		parts = append(parts, buildPageBreakParagraph())
	}
	parts = append(parts, doc.BodyContent)
	entries["word/document.xml"] = []byte(doc.Prefix + strings.Join(parts, "") + doc.Suffix + "</w:document>")

	if err := ensureDocxUpdateFields(entries); err != nil {
		return err
	}

	if err := os.MkdirAll(filepath.Dir(outputPath), 0755); err != nil && filepath.Dir(outputPath) != "." {
		return fmt.Errorf("create output directory: %w", err)
	}
	return writeDocxEntries(outputPath, entries)
}

func ensureDocxUpdateFields(entries map[string][]byte) error {
	if err := ensureSettingsXML(entries); err != nil {
		return err
	}
	if err := ensureSettingsRelationship(entries); err != nil {
		return err
	}
	if err := ensureSettingsContentType(entries); err != nil {
		return err
	}
	return nil
}

func ensureSettingsXML(entries map[string][]byte) error {
	settingsXML, ok := entries[settingsPath]
	if !ok || len(bytes.TrimSpace(settingsXML)) == 0 {
		entries[settingsPath] = []byte(defaultSettingsXML)
		return nil
	}

	text := string(settingsXML)
	if updateFieldsPattern.MatchString(text) {
		text = updateFieldsPattern.ReplaceAllString(text, updateFieldsElement)
		entries[settingsPath] = []byte(text)
		return nil
	}

	if settingsSelfPattern.MatchString(text) {
		text = settingsSelfPattern.ReplaceAllString(text, `<w:settings$1>`+updateFieldsElement+`</w:settings>`)
		entries[settingsPath] = []byte(text)
		return nil
	}

	updated, err := insertXMLBeforeClosingTag(text, "</w:settings>", updateFieldsElement)
	if err != nil {
		return fmt.Errorf("update %s: %w", settingsPath, err)
	}
	entries[settingsPath] = []byte(updated)
	return nil
}

func ensureSettingsRelationship(entries map[string][]byte) error {
	relsXML := string(entries[documentRelsPath])
	if strings.TrimSpace(relsXML) == "" {
		relsXML = defaultRelationshipsXML
	}

	var settingsRelID string
	for _, tag := range relationshipTagPattern.FindAllString(relsXML, -1) {
		if !strings.Contains(tag, `Type="`+settingsRelationshipType+`"`) {
			continue
		}
		settingsRelID = xmlAttrValue(tag, "Id")
		if settingsRelID == "" {
			settingsRelID = nextRelationshipID(relsXML)
		}
		replacement := fmt.Sprintf(`<Relationship Id="%s" Type="%s" Target="settings.xml"/>`, settingsRelID, settingsRelationshipType)
		relsXML = strings.Replace(relsXML, tag, replacement, 1)
		entries[documentRelsPath] = []byte(relsXML)
		return nil
	}

	settingsRelID = nextRelationshipID(relsXML)
	relElement := fmt.Sprintf(`<Relationship Id="%s" Type="%s" Target="settings.xml"/>`, settingsRelID, settingsRelationshipType)

	if strings.Contains(relsXML, "/>") && strings.Contains(relsXML, "<Relationships") && strings.Contains(relsXML, "</Relationships>") == false {
		relsXML = strings.Replace(relsXML, "/>", ">"+relElement+"</Relationships>", 1)
		entries[documentRelsPath] = []byte(relsXML)
		return nil
	}

	updated, err := insertXMLBeforeClosingTag(relsXML, "</Relationships>", relElement)
	if err != nil {
		return fmt.Errorf("update %s: %w", documentRelsPath, err)
	}
	entries[documentRelsPath] = []byte(updated)
	return nil
}

func ensureSettingsContentType(entries map[string][]byte) error {
	contentTypesXML, ok := entries[contentTypesPath]
	if !ok || len(bytes.TrimSpace(contentTypesXML)) == 0 {
		return fmt.Errorf("missing %s", contentTypesPath)
	}

	text := string(contentTypesXML)
	for _, tag := range overrideTagPattern.FindAllString(text, -1) {
		if !strings.Contains(tag, `PartName="/word/settings.xml"`) {
			continue
		}
		replacement := defaultSettingsOverride
		text = strings.Replace(text, tag, replacement, 1)
		entries[contentTypesPath] = []byte(text)
		return nil
	}

	updated, err := insertXMLBeforeClosingTag(text, "</Types>", defaultSettingsOverride)
	if err != nil {
		return fmt.Errorf("update %s: %w", contentTypesPath, err)
	}
	entries[contentTypesPath] = []byte(updated)
	return nil
}

func readDocxEntries(path string) (map[string][]byte, error) {
	if strings.ToLower(filepath.Ext(path)) != ".docx" {
		return nil, fmt.Errorf("input must be a .docx file: %s", path)
	}

	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("open DOCX %s: %w", path, err)
	}
	defer reader.Close()

	entries := make(map[string][]byte, len(reader.File))
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return nil, fmt.Errorf("open zip entry %s: %w", file.Name, err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, fmt.Errorf("read zip entry %s: %w", file.Name, err)
		}
		entries[file.Name] = data
	}

	if len(entries["word/document.xml"]) == 0 {
		return nil, fmt.Errorf("missing word/document.xml in %s", path)
	}
	return entries, nil
}

func writeDocxEntries(outputPath string, entries map[string][]byte) error {
	file, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("create output DOCX: %w", err)
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	names := make([]string, 0, len(entries))
	for name := range entries {
		names = append(names, name)
	}
	sort.Strings(names)

	for _, name := range names {
		entryWriter, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			return fmt.Errorf("create zip entry %s: %w", name, err)
		}
		if _, err := entryWriter.Write(entries[name]); err != nil {
			_ = writer.Close()
			return fmt.Errorf("write zip entry %s: %w", name, err)
		}
	}

	if err := writer.Close(); err != nil {
		return fmt.Errorf("close DOCX writer: %w", err)
	}
	return nil
}

func parseWordDocument(docXML []byte) (parsedWordDocument, error) {
	match := bodyPattern.FindSubmatch(docXML)
	if len(match) != 4 {
		return parsedWordDocument{}, fmt.Errorf("document.xml does not contain a parseable w:body")
	}

	return parsedWordDocument{
		Prefix:      string(match[1]),
		BodyContent: string(match[2]),
		Suffix:      string(match[3]),
	}, nil
}

func buildTextParagraph(text string) string {
	return `<w:p><w:r><w:t xml:space="preserve">` + xmlEscape(text) + `</w:t></w:r></w:p>`
}

func buildTOCFieldParagraph(fieldCode string) string {
	return `<w:p>` +
		`<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
		`<w:r><w:instrText xml:space="preserve"> ` + xmlEscape(fieldCode) + ` </w:instrText></w:r>` +
		`<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
		`<w:r><w:t>Update the TOC in Word to calculate page numbers.</w:t></w:r>` +
		`<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
		`</w:p>`
}

func buildPageBreakParagraph() string {
	return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`
}

func buildTOCSectionBreakParagraphWithoutHeaderFooter(templateDocumentXML []byte) string {
	sectPr := extractBodyLevelSectPr(templateDocumentXML)
	if strings.TrimSpace(sectPr) == "" {
		return ""
	}

	// TOC 首页作为独立 section：去掉页眉页脚引用，避免目录页显示页眉页脚。
	sectPr = headerFooterRefPattern.ReplaceAllString(sectPr, "")
	return `<w:p><w:pPr>` + sectPr + `</w:pPr></w:p>`
}

func extractBodyLevelSectPr(documentXML []byte) string {
	doc := string(documentXML)
	bodyStart := strings.Index(doc, "<w:body")
	if bodyStart < 0 {
		return ""
	}
	bodyOpenEnd := strings.Index(doc[bodyStart:], ">")
	if bodyOpenEnd < 0 {
		return ""
	}
	bodyOpenEnd += bodyStart

	bodyClose := strings.Index(doc[bodyOpenEnd+1:], "</w:body>")
	if bodyClose < 0 {
		return ""
	}
	bodyClose += bodyOpenEnd + 1

	bodyContent := doc[bodyOpenEnd+1 : bodyClose]
	sectStart := strings.LastIndex(bodyContent, "<w:sectPr")
	if sectStart < 0 {
		return ""
	}

	candidate := bodyContent[sectStart:]
	match := sectPrSelfPattern.FindString(candidate)
	if match == "" {
		match = sectPrFullPattern.FindString(candidate)
	}
	if match == "" {
		return ""
	}

	remaining := strings.TrimSpace(candidate[len(match):])
	if remaining != "" {
		return ""
	}
	return match
}

func xmlEscape(value string) string {
	replacer := strings.NewReplacer(
		"&", "&amp;",
		"<", "&lt;",
		">", "&gt;",
		`"`, "&quot;",
		"'", "&apos;",
	)
	return replacer.Replace(value)
}

func insertXMLBeforeClosingTag(xmlText, closingTag, insertion string) (string, error) {
	idx := strings.Index(xmlText, closingTag)
	if idx < 0 {
		return "", fmt.Errorf("closing tag %s not found", closingTag)
	}
	return xmlText[:idx] + insertion + xmlText[idx:], nil
}

func xmlAttrValue(tag, attr string) string {
	pattern := regexp.MustCompile(fmt.Sprintf(`\b%s="([^"]*)"`, regexp.QuoteMeta(attr)))
	match := pattern.FindStringSubmatch(tag)
	if len(match) != 2 {
		return ""
	}
	return match[1]
}

func nextRelationshipID(relsXML string) string {
	maxID := 0
	for _, match := range rIDPattern.FindAllStringSubmatch(relsXML, -1) {
		if len(match) != 2 {
			continue
		}
		id, err := strconv.Atoi(match[1])
		if err != nil {
			continue
		}
		if id > maxID {
			maxID = id
		}
	}
	return fmt.Sprintf("rId%d", maxID+1)
}

func CombineDocx(srcPath []string, outPath string, outFile string, generateTOC string, logCallback LogCallback) error {
	logCallback("🔍 Terminating all Word processes...\n")
	_ = KillWordProcesses()

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	defer ole.CoUninitialize()

	var files []string
	for _, p := range srcPath {
		if !strings.HasPrefix(filepath.Base(p), "~$") &&
			(strings.HasSuffix(strings.ToLower(p), ".docx") || strings.HasSuffix(strings.ToLower(p), ".rtf")) {
			files = append(files, p)
		}
	}

	if len(files) == 0 {
		logCallback("❌ No files to merge found.\n")
		return fmt.Errorf("No files to merge found.")
	}

	logCallback("⏳ Starting Word Application...\n")
	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		logCallback("❌ Failed to create WORD object: %v\n", err)
		return err
	}
	word, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer word.Release()

	oleutil.PutProperty(word, "Visible", false)
	oleutil.PutProperty(word, "DisplayAlerts", 0)

	documents := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
	defer documents.Release()

	absFirstPath, _ := filepath.Abs(files[0])
	logCallback("📖 Opening base document: %s\n", filepath.Base(absFirstPath))
	mainDoc := oleutil.MustCallMethod(documents, "Open", absFirstPath).ToIDispatch()

	logCallback("🔄 Merging %d documents...\n", len(files)-1)
	for i := 1; i < len(files); i++ {
		absPath, _ := filepath.Abs(files[i])
		logCallback("  -> Merging: %s\n", filepath.Base(absPath))
		content := oleutil.MustGetProperty(mainDoc, "Content").ToIDispatch()

		oleutil.MustCallMethod(content, "Collapse", wdCollapseEnd)
		oleutil.MustCallMethod(content, "InsertBreak", wdSectionBreakNextPage)
		oleutil.MustCallMethod(content, "Collapse", wdCollapseEnd)
		oleutil.MustCallMethod(content, "InsertFile", absPath)
		content.Release()
	}

	_ = os.MkdirAll(outPath, 0755)
	tempRtfPath, _ := filepath.Abs(filepath.Join(outPath, outFile+".rtf"))

	logCallback("💾 Saving as temporary RTF...\n")
	oleutil.MustCallMethod(mainDoc, "SaveAs2", tempRtfPath, wdFormatRTF)
	oleutil.MustCallMethod(mainDoc, "Close")
	mainDoc.Release()

	logCallback("⚙️  Processing \\pgnrestart in RTF...\n")
	// 处理 \pgnrestart
	rtfBytes, _ := os.ReadFile(tempRtfPath)
	keyword := []byte("\\pgnrestart")
	if parts := bytes.Split(rtfBytes, keyword); len(parts) > 2 {
		var buf bytes.Buffer
		buf.Write(parts[0])
		buf.Write(keyword)
		for i := 1; i < len(parts); i++ {
			buf.Write(parts[i])
		}
		_ = os.WriteFile(tempRtfPath, buf.Bytes(), 0644)
	}

	logCallback("📝 Converting RTF back to Docx...\n")
	finalDoc := oleutil.MustCallMethod(documents, "Open", tempRtfPath).ToIDispatch()
	finalDocxPath, _ := filepath.Abs(filepath.Join(outPath, outFile+".docx"))
	oleutil.MustCallMethod(finalDoc, "SaveAs2", finalDocxPath, wdFormatDocumentDefault)
	oleutil.MustCallMethod(finalDoc, "Close")
	finalDoc.Release()

	needTOC := strings.EqualFold(strings.TrimSpace(generateTOC), "Y")
	if needTOC {
		logCallback("📚 Injecting TOC into final Docx (OpenXML)...\n")
		if err := AddTOCToDocx(finalDocxPath, finalDocxPath); err != nil {
			logCallback("❌ Failed to inject TOC: %v\n", err)
			oleutil.MustCallMethod(word, "Quit")
			return err
		}

		logCallback("🌟 Reopening Document via COM to force Update TOC...\n")
		updatedDocVar, err := oleutil.CallMethod(documents, "Open", finalDocxPath)
		if err != nil {
			logCallback("❌ Failed to reopen DOCX after TOC injection: %v\n", err)
			return fmt.Errorf("failed to reopen DOCX after TOC injection: %w", err)
		}
		updatedDoc := updatedDocVar.ToIDispatch()
		defer updatedDoc.Release()

		tocsVar, err := oleutil.GetProperty(updatedDoc, "TablesOfContents")
		if err != nil {
			logCallback("⚠️ Warning: Unable to access TablesOfContents: %v\n", err)
		} else {
			tocs := tocsVar.ToIDispatch()
			countVar, err := oleutil.GetProperty(tocs, "Count")
			if err != nil {
				logCallback("⚠️ Warning: Unable to read TOC count: %v\n", err)
			} else {
				tocCount := int(countVar.Val)
				if tocCount > 0 {
					logCallback("✨ Found %d TOC(s). Updating fields and page numbers...\n", tocCount)
					for i := 1; i <= tocCount; i++ {
						tocItemVar, itemErr := oleutil.CallMethod(tocs, "Item", i)
						if itemErr != nil {
							logCallback("⚠️ Warning: Failed to get TOC item %d: %v\n", i, itemErr)
							continue
						}
						tocItem := tocItemVar.ToIDispatch()
						if _, updateErr := oleutil.CallMethod(tocItem, "Update"); updateErr != nil {
							logCallback("⚠️ Warning: Failed to update TOC item %d: %v\n", i, updateErr)
						}
						tocItem.Release()
					}
				} else {
					logCallback("⚠️ Warning: No TOC found by Word COM.\n")
				}
			}
			tocs.Release()
		}

		if _, err := oleutil.CallMethod(updatedDoc, "Save"); err != nil {
			logCallback("⚠️ Warning: Failed to save DOCX after TOC update: %v\n", err)
		}
		if _, err := oleutil.CallMethod(updatedDoc, "Close"); err != nil {
			logCallback("⚠️ Warning: Failed to close DOCX after TOC update: %v\n", err)
		}
	} else {
		logCallback("ℹ️ TOC generation skipped (generateTOC != Y).\n")
	}

	logCallback("🚪 Quitting Word Application...\n")
	oleutil.MustCallMethod(word, "Quit")
	_ = os.Remove(tempRtfPath)

	logCallback("🎉 Task completed successfully.\n")
	return nil
}

// ==============================================================================
// 8. File Scanning & Sorting (extracted from GUI logic for CLI use)
// ==============================================================================

type SortableFile struct {
	Name     string
	FullPath string
	Ord      int
	SortNums []int
}

// ScanAndSortFiles scans a directory for files with given extensions, then sorts
// them using the same T/F/L prefix and numeric ordering logic as the GUI version.
func ScanAndSortFiles(dirPath string, allowedExts []string) ([]string, error) {
	entries, err := os.ReadDir(dirPath)
	if err != nil {
		return nil, fmt.Errorf("failed to read directory %s: %w", dirPath, err)
	}

	var sortable []SortableFile
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		nameLower := strings.ToLower(name)

		if strings.HasPrefix(name, "~") {
			continue
		}

		isValidExt := false
		for _, ext := range allowedExts {
			if strings.HasSuffix(nameLower, strings.ToLower(ext)) {
				isValidExt = true
				break
			}
		}
		if !isValidExt {
			continue
		}

		sf := SortableFile{
			Name:     name,
			FullPath: filepath.Join(dirPath, name),
		}

		if strings.HasPrefix(nameLower, "t") {
			sf.Ord = 1
		} else if strings.HasPrefix(nameLower, "f") {
			sf.Ord = 2
		} else if strings.HasPrefix(nameLower, "l") {
			sf.Ord = 3
		} else {
			sf.Ord = 999
		}

		re := regexp.MustCompile(`^[tTfFlL][-.]`)
		name1 := re.ReplaceAllString(name, "")
		name1 = regexp.MustCompile(`[a-zA-Z]`).ReplaceAllString(name1, "")
		name1 = regexp.MustCompile(`^-+`).ReplaceAllString(name1, "")
		name1 = regexp.MustCompile(`-+`).ReplaceAllString(name1, "-")
		name1 = strings.ReplaceAll(name1, ".", "-")
		name1 = regexp.MustCompile(`(\d)[^\d]*$`).ReplaceAllString(name1, "$1")

		parts := strings.Split(name1, "-")
		for _, part := range parts {
			if part == "" {
				continue
			}
			num, err := strconv.Atoi(part)
			if err != nil {
				num = 9999
			}
			sf.SortNums = append(sf.SortNums, num)
		}

		sortable = append(sortable, sf)
	}

	sort.Slice(sortable, func(i, j int) bool {
		if sortable[i].Ord != sortable[j].Ord {
			return sortable[i].Ord < sortable[j].Ord
		}
		maxLen := len(sortable[i].SortNums)
		if len(sortable[j].SortNums) > maxLen {
			maxLen = len(sortable[j].SortNums)
		}
		for k := 0; k < maxLen; k++ {
			var numI, numJ int
			if k < len(sortable[i].SortNums) {
				numI = sortable[i].SortNums[k]
			}
			if k < len(sortable[j].SortNums) {
				numJ = sortable[j].SortNums[k]
			}
			if numI != numJ {
				return numI < numJ
			}
		}
		return sortable[i].Name < sortable[j].Name
	})

	var result []string
	for _, sf := range sortable {
		result = append(result, sf.FullPath)
	}
	return result, nil
}
