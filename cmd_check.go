package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var checkDir string

var checkCmd = &cobra.Command{
	Use:   "check",
	Short: "Validate RTF page-count consistency across a folder",
	Long: `Validates page-count consistency for all RTF files in the specified folder
(including subfolders). Compares the actual page count rendered by Word (via COM)
against the "Page 1 of N" marker parsed from the RTF text.

Requires Microsoft Word to be installed. Close all open Word documents before running.`,
	Example: `  rtftool check -d "C:\Projects\RTF_Output"
  rtftool check --dir "D:\reports\rtf_files"`,
	RunE: func(cmd *cobra.Command, args []string) error {
		if checkDir == "" {
			return fmt.Errorf("--dir is required: specify the RTF folder path")
		}

		if info, err := os.Stat(checkDir); err != nil || !info.IsDir() {
			return fmt.Errorf("invalid directory path: %s", checkDir)
		}

		logCallback := func(format string, args ...interface{}) {
			fmt.Printf(format, args...)
		}

		result := RTFPageCheck(checkDir, logCallback)

		fmt.Println("\n========================================")
		fmt.Printf("Total: %d | Matched: %d | Mismatched: %d | Failed: %d\n",
			result.TotalFiles, result.SuccessCount,
			result.TotalFiles-result.SuccessCount-result.FailedCount,
			result.FailedCount)
		fmt.Printf("Duration: %.1fs\n", result.Duration.Seconds())

		if result.Error != "" {
			return fmt.Errorf("check failed: %s", result.Error)
		}
		if !result.AllMatched {
			os.Exit(2)
		}
		return nil
	},
}

func init() {
	checkCmd.Flags().StringVarP(&checkDir, "dir", "d", "", "RTF folder path to check (required)")
	_ = checkCmd.MarkFlagRequired("dir")
	rootCmd.AddCommand(checkCmd)
}
