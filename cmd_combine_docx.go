package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
)

var (
	combineDocxOutDir   string
	combineDocxOutName  string
	combineDocxInputDir string
	combineDocxSort     bool
)

var combineDocxCmd = &cobra.Command{
	Use:   "combine-docx [flags] [file1 file2 ...]",
	Short: "Merge multiple DOCX/RTF files into one DOCX (General style)",
	Long: `Merges multiple DOCX or RTF files into a single DOCX document via Word COM.
Does not support Table of Contents (TOC) insertion.

Requires Microsoft Word to be installed. Close all open Word documents before running.

File source (choose one):
  1. List file paths as positional arguments (merged in the given order)
  2. Use --input-dir to scan a folder for .docx/.rtf files (auto-sorted by default)`,
	Example: `  rtftool combine-docx -o "C:\Output" -n "Merged" "C:\docs\part1.docx" "C:\docs\part2.docx"
  rtftool combine-docx -i "C:\docs\source" -o "C:\Output" -n "Merged"`,
	RunE: func(cmd *cobra.Command, args []string) error {
		if combineDocxOutDir == "" {
			return fmt.Errorf("--out-dir is required")
		}
		if combineDocxOutName == "" {
			return fmt.Errorf("--out-name is required")
		}

		var filePaths []string

		if combineDocxInputDir != "" {
			if info, err := os.Stat(combineDocxInputDir); err != nil || !info.IsDir() {
				return fmt.Errorf("invalid input directory: %s", combineDocxInputDir)
			}
			if combineDocxSort {
				sorted, err := ScanAndSortFiles(combineDocxInputDir, []string{".docx", ".rtf"})
				if err != nil {
					return err
				}
				filePaths = sorted
			} else {
				entries, err := os.ReadDir(combineDocxInputDir)
				if err != nil {
					return err
				}
				for _, e := range entries {
					if e.IsDir() || strings.HasPrefix(e.Name(), "~") {
						continue
					}
					lower := strings.ToLower(e.Name())
					if strings.HasSuffix(lower, ".docx") || strings.HasSuffix(lower, ".rtf") {
						filePaths = append(filePaths, filepath.Join(combineDocxInputDir, e.Name()))
					}
				}
			}
		} else if len(args) > 0 {
			filePaths = args
		}

		if len(filePaths) == 0 {
			return fmt.Errorf("no files specified. Use positional args or --input-dir")
		}

		fmt.Printf("[INFO] Merging %d files...\n", len(filePaths))
		for i, p := range filePaths {
			fmt.Printf("  %d. %s\n", i+1, p)
		}
		fmt.Printf("[INFO] Output: %s\n", filepath.Join(combineDocxOutDir, combineDocxOutName+".docx"))
		fmt.Println("----------------------------------------")

		logCallback := func(format string, args ...interface{}) {
			fmt.Printf(format, args...)
		}

		err := CombineDocx(filePaths, combineDocxOutDir, combineDocxOutName, logCallback)
		if err != nil {
			return fmt.Errorf("combine-docx failed: %v", err)
		}

		fmt.Println("[INFO] Combine-docx completed successfully!")
		return nil
	},
}

func init() {
	combineDocxCmd.Flags().StringVarP(&combineDocxOutDir, "out-dir", "o", "", "Output directory (required)")
	combineDocxCmd.Flags().StringVarP(&combineDocxOutName, "out-name", "n", "", "Output file name without extension (required)")
	combineDocxCmd.Flags().StringVarP(&combineDocxInputDir, "input-dir", "i", "", "Scan files from this directory")
	combineDocxCmd.Flags().BoolVarP(&combineDocxSort, "sort", "s", true, "Auto-sort files by T/F/L prefix and numbers")

	_ = combineDocxCmd.MarkFlagRequired("out-dir")
	_ = combineDocxCmd.MarkFlagRequired("out-name")

	rootCmd.AddCommand(combineDocxCmd)
}
