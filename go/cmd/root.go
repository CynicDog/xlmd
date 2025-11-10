package cmd

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"xlmd/pkg/excel"
	"xlmd/pkg/markdown"
)

type Options struct {
	InputFile  string
	OutputFile string
	Direction  string // "excel2md" or "md2excel"
}

// ParseFlags parses CLI args using -i and -o and infers conversion direction
func ParseFlags() Options {
	var opts Options

	flag.StringVar(&opts.InputFile, "i", "", "Input file path (Excel or Markdown)")
	flag.StringVar(&opts.OutputFile, "o", "", "Output file path")
	flag.Parse()

	if opts.InputFile == "" || opts.OutputFile == "" {
		flag.Usage()
		os.Exit(1)
	}

	inExt := filepath.Ext(opts.InputFile)
	outExt := filepath.Ext(opts.OutputFile)

	switch {
	case inExt == ".xlsx" && outExt == ".md":
		opts.Direction = "excel2md"
	case inExt == ".md" && outExt == ".xlsx":
		opts.Direction = "md2excel"
	default:
		log.Fatalf("Invalid file combination. Must be either Excel → Markdown (.xlsx → .md) or Markdown → Excel (.md → .xlsx)")
	}

	return opts
}

func Run() {
	opts := ParseFlags()

	const banner = `
 __  __     __         __    __     _____    
/\_\_\_\   /\ \       /\ "-./  \   /\  __-.  
\/_/\_\/_  \ \ \____  \ \ \-./\ \  \ \ \/\ \ 
  /\_\/\_\  \ \_____\  \ \_\ \ \_\  \ \____- 
  \/_/\/_/   \/_____/   \/_/  \/_/   \/____/ 
  
maintained at https://github.com/CynicDog/xlmd 
`

	fmt.Printf("%s\n\nConvert %s input file to %s output file\n\n", banner, opts.InputFile, opts.OutputFile)

	switch opts.Direction {
	case "excel2md":
		data, err := excel.ReadExcel(opts.InputFile)
		if err != nil {
			log.Fatalf("Failed to read Excel: %v", err)
		}
		md := markdown.ToMarkdown(data)
		if err := os.WriteFile(opts.OutputFile, []byte(md), 0644); err != nil {
			log.Fatalf("Failed to write Markdown: %v", err)
		}

	case "md2excel":
		data, err := markdown.ReadMarkdown(opts.InputFile)
		if err != nil {
			log.Fatalf("Failed to read Markdown: %v", err)
		}
		if err := excel.WriteExcel(opts.OutputFile, data); err != nil {
			log.Fatalf("Failed to write Excel: %v", err)
		}
	}
}
