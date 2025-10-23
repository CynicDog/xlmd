package markdown

import (
	"strings"
	"xlmd/pkg/excel"
)

// ToMarkdown converts SheetData into a Markdown document string
func ToMarkdown(sheets []excel.SheetData) string {
	var sb strings.Builder

	for _, sheet := range sheets {
		// skip empty sheets
		if len(sheet.Rows) == 0 {
			continue
		}

		// section heading
		sb.WriteString("## " + sheet.Name + "\n\n")

		// determine column count (based on longest row)
		colCount := 0
		for _, r := range sheet.Rows {
			if len(r) > colCount {
				colCount = len(r)
			}
		}

		// pad rows to equal length
		for i, row := range sheet.Rows {
			if len(row) < colCount {
				padded := make([]string, colCount)
				copy(padded, row)
				for j := len(row); j < colCount; j++ {
					padded[j] = ""
				}
				sheet.Rows[i] = padded
			}
		}

		// first row as header
		header := sheet.Rows[0]
		sb.WriteString("| " + strings.Join(header, " | ") + " |\n")

		// separator row
		sb.WriteString("|" + strings.Repeat(" --- |", colCount) + "\n")

		// remaining rows
		for _, row := range sheet.Rows[1:] {
			sb.WriteString("| " + strings.Join(row, " | ") + " |\n")
		}

		sb.WriteString("\n")
	}

	return sb.String()
}
