package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
)

var (
	convertToPdf  bool
	convertToDocx bool
)

var convertCmd = &cobra.Command{
	Use:   "convert [flags] <source.rtf>",
	Short: "Convert RTF to PDF and/or DOCX",
	Long: `Converts a single RTF file to PDF and/or DOCX format via Word COM.
PDF output is automatically optimized (bookmark expansion + Fast Web View).

Requires Microsoft Word to be installed. Close all open Word documents before running.
At least one of --pdf or --docx must be specified.`,
	Example: `  rtftool convert --pdf --docx "C:\reports\report.rtf"
  rtftool convert --pdf "D:\output\combined.rtf"`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		if !convertToPdf && !convertToDocx {
			return fmt.Errorf("at least one of --pdf or --docx must be specified")
		}

		rtfFile := args[0]
		if !strings.HasSuffix(strings.ToLower(rtfFile), ".rtf") {
			return fmt.Errorf("source file must be an .rtf file: %s", rtfFile)
		}
		if _, err := os.Stat(rtfFile); err != nil {
			return fmt.Errorf("source file not found: %s", rtfFile)
		}

		fmt.Printf("[INFO] Source: %s\n", rtfFile)
		fmt.Printf("[INFO] Convert to PDF: %v | Convert to DOCX: %v\n", convertToPdf, convertToDocx)
		fmt.Println("----------------------------------------")

		logCallback := func(format string, args ...interface{}) {
			fmt.Printf(format, args...)
		}

		err := RTFConverter(rtfFile, convertToPdf, convertToDocx, logCallback)
		if err != nil {
			return fmt.Errorf("conversion failed: %v", err)
		}

		fmt.Println("[INFO] Conversion completed successfully!")
		return nil
	},
}

func init() {
	convertCmd.Flags().BoolVar(&convertToPdf, "pdf", false, "Convert to PDF")
	convertCmd.Flags().BoolVar(&convertToDocx, "docx", false, "Convert to DOCX")

	rootCmd.AddCommand(convertCmd)
}
