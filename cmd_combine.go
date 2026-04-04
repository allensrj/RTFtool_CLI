package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
)

var (
	combineOutDir      string
	combineOutName     string
	combineToc         bool
	combineTocRows     int
	combineRefreshPage bool
	combineInputDir    string
	combineSort        bool
)

var combineCmd = &cobra.Command{
	Use:   "combine [flags] [file1.rtf file2.rtf ...]",
	Short: "Merge multiple RTF files into one (Specify style)",
	Long: `Merges multiple RTF files into a single RTF document using Specify style.
Supports automatic Table of Contents (TOC) generation and global page-number refresh.

This command does NOT require Word COM — it operates on RTF text directly.

File source (choose one):
  1. List RTF file paths as positional arguments (merged in the given order)
  2. Use --input-dir to scan a folder for RTF files (auto-sorted by default)`,
	Example: `  # Manual file list
  rtftool combine -o "C:\Output" -n "Combined" -t -p "C:\RTF\T-14-1.rtf" "C:\RTF\F-14-1.rtf"

  # Scan folder with auto-sort
  rtftool combine -i "C:\RTF\source" -o "C:\Output" -n "Combined" --toc --refresh-page`,
	RunE: func(cmd *cobra.Command, args []string) error {
		if combineOutDir == "" {
			return fmt.Errorf("--out-dir is required")
		}
		if combineOutName == "" {
			return fmt.Errorf("--out-name is required")
		}

		var filePaths []string

		if combineInputDir != "" {
			if info, err := os.Stat(combineInputDir); err != nil || !info.IsDir() {
				return fmt.Errorf("invalid input directory: %s", combineInputDir)
			}
			if combineSort {
				sorted, err := ScanAndSortFiles(combineInputDir, []string{".rtf"})
				if err != nil {
					return err
				}
				filePaths = sorted
			} else {
				entries, err := os.ReadDir(combineInputDir)
				if err != nil {
					return err
				}
				for _, e := range entries {
					if !e.IsDir() && strings.HasSuffix(strings.ToLower(e.Name()), ".rtf") && !strings.HasPrefix(e.Name(), "~") {
						filePaths = append(filePaths, filepath.Join(combineInputDir, e.Name()))
					}
				}
			}
		} else if len(args) > 0 {
			filePaths = args
		}

		if len(filePaths) == 0 {
			return fmt.Errorf("no RTF files specified. Use positional args or --input-dir")
		}

		fmt.Printf("[INFO] Merging %d RTF files...\n", len(filePaths))
		for i, p := range filePaths {
			fmt.Printf("  %d. %s\n", i+1, p)
		}
		fmt.Printf("[INFO] TOC: %v | TOC Rows: %d | Refresh Page: %v\n", combineToc, combineTocRows, combineRefreshPage)
		fmt.Printf("[INFO] Output: %s\n", filepath.Join(combineOutDir, combineOutName+".rtf"))
		fmt.Println("----------------------------------------")

		if err := os.MkdirAll(combineOutDir, 0755); err != nil {
			return fmt.Errorf("failed to create output directory: %v", err)
		}

		err := CombineRTF(filePaths, combineToc, combineTocRows, combineRefreshPage, combineOutDir, combineOutName)
		if err != nil {
			return fmt.Errorf("combine failed: %v", err)
		}

		fmt.Println("[INFO] Combine completed successfully!")
		return nil
	},
}

func init() {
	combineCmd.Flags().StringVarP(&combineOutDir, "out-dir", "o", "", "Output directory (required)")
	combineCmd.Flags().StringVarP(&combineOutName, "out-name", "n", "", "Output file name without extension (required)")
	combineCmd.Flags().BoolVarP(&combineToc, "toc", "t", false, "Add Table of Contents")
	combineCmd.Flags().IntVar(&combineTocRows, "toc-rows", 23, "TOC rows per page")
	combineCmd.Flags().BoolVarP(&combineRefreshPage, "refresh-page", "p", false, "Refresh page numbers")
	combineCmd.Flags().StringVarP(&combineInputDir, "input-dir", "i", "", "Scan RTF files from this directory")
	combineCmd.Flags().BoolVarP(&combineSort, "sort", "s", true, "Auto-sort files by T/F/L prefix and numbers")

	_ = combineCmd.MarkFlagRequired("out-dir")
	_ = combineCmd.MarkFlagRequired("out-name")

	rootCmd.AddCommand(combineCmd)
}
