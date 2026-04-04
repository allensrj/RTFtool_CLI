package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var docx2rtfCmd = &cobra.Command{
	Use:   "docx2rtf <path>",
	Short: "Convert DOCX file(s) to RTF",
	Long: `Converts a single .docx file or all .docx files in a folder to RTF format.
Existing RTF files with the same name are automatically skipped.

Requires Microsoft Word to be installed. Close all open Word documents before running.`,
	Example: `  # Single file
  rtftool docx2rtf "C:\docs\report.docx"

  # Batch conversion (all .docx in folder)
  rtftool docx2rtf "C:\docs\batch_folder"`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		inputPath := args[0]
		if _, err := os.Stat(inputPath); err != nil {
			return fmt.Errorf("path not found: %s", inputPath)
		}

		fmt.Printf("[INFO] Target path: %s\n", inputPath)
		fmt.Println("----------------------------------------")

		logCallback := func(format string, args ...interface{}) {
			fmt.Printf(format, args...)
		}

		result := ConvertDocxToRTF(inputPath, logCallback)

		fmt.Println("\n========================================")
		fmt.Printf("Total: %d | Success: %d | Failed: %d\n",
			result.TotalFiles, result.SuccessCount, result.ErrorCount)
		fmt.Printf("Duration: %v\n", result.Duration)

		if result.Error != "" && result.Error != "no files found" {
			return fmt.Errorf("conversion failed: %s", result.Error)
		}
		if result.ErrorCount > 0 {
			os.Exit(2)
		}
		return nil
	},
}

func init() {
	rootCmd.AddCommand(docx2rtfCmd)
}
