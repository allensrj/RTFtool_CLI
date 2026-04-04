package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "rtftool",
	Short: "RTF Tools CLI - RTF/DOCX file processing toolkit",
	Long: `RTF Tools CLI v0.3

A command-line toolkit for RTF and DOCX file operations:
  check        Validate RTF page-count consistency
  combine      Merge multiple RTF files (Specify style, with TOC support)
  combine-docx Merge multiple DOCX/RTF files (General style)
  convert      Convert RTF to PDF/DOCX
  docx2rtf     Convert DOCX to RTF

Prerequisites: Windows with Microsoft Word installed.
Close all open Word documents before running any command.`,
}

func main() {
	rootCmd.CompletionOptions.DisableDefaultCmd = true
	rootCmd.SetHelpTemplate(rootCmd.HelpTemplate() + fmt.Sprintf("\nVersion: %s\n", "0.3.0"))

	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}
