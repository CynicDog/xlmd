package markdown

import (
	"fmt"
	"os"
	"regexp"
	"strings"
	"xlmd/pkg/excel"
)

// ReadMarkdown reads a Markdown file and returns all tables as SheetData
// This implementation uses standard library functions to parse the Markdown table format.
func ReadMarkdown(filePath string) ([]excel.SheetData, error) {
	content, err := os.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to read markdown file: %w", err)
	}

	// Use regex to find all "## Sheet Name\n\n" occurrences
	// (?m) enables multiline mode for ^ (start of line)
	// ([^\n]+) captures the sheet name
	re := regexp.MustCompile(`(?m)^## ([^\n]+)\n\n`)

	// Get all sheet names and their surrounding content blocks
	matches := re.FindAllStringSubmatch(string(content), -1)

	// Split the content based on the regex matches.
	// The first part (parts[0]) is pre-sheet content, which is ignored.
	parts := re.Split(string(content), -1)

	sheets := make([]excel.SheetData, 0, len(matches))

	if len(matches) == 0 {
		// If no '## Sheet Name' headings are found, treat the entire file as a single default sheet
		tableData, err := parseTable(string(content))
		if err != nil {
			return nil, err
		}
		if len(tableData) > 0 {
			sheets = append(sheets, excel.SheetData{
				Name: "Sheet1", // Default name for un-named sheets
				Rows: tableData,
			})
		}
		return sheets, nil
	}

	// Process content block after each sheet header
	for i, match := range matches {
		sheetName := strings.TrimSpace(match[1])
		// parts[i+1] corresponds to the content following match[i]
		sheetContent := strings.TrimSpace(parts[i+1])

		if sheetContent == "" {
			continue
		}

		tableData, err := parseTable(sheetContent)
		if err != nil {
			return nil, fmt.Errorf("error parsing table for sheet %s: %w", sheetName, err)
		}

		if len(tableData) > 0 {
			// Markdown sheet names can't contain forward slashes, but Excel sheet names can't either,
			// so we ensure basic sanitization.
			sheetName = strings.ReplaceAll(sheetName, "/", "-")

			sheets = append(sheets, excel.SheetData{
				Name: sheetName,
				Rows: tableData,
			})
		}
	}

	return sheets, nil
}

// parseTable takes a string containing a Markdown table and extracts the rows and cells.
func parseTable(mdContent string) ([][]string, error) {
	lines := strings.Split(mdContent, "\n")
	rows := [][]string{}

	expectedCols := 0

	// Regex to check if a string contains only separator characters (dash, colon, space)
	// and must contain at least one dash to distinguish it from an empty cell or pure space.
	separatorRe := regexp.MustCompile(`^[:\s-]*$`)

	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}

		// Check if it's a pipe-delimited row
		if strings.HasPrefix(line, "|") && strings.HasSuffix(line, "|") {

			// Remove leading and trailing pipe and trim
			trimmed := strings.TrimSpace(line[1 : len(line)-1])

			// Split by '|' to examine the content of each "cell"
			rawCells := strings.Split(trimmed, "|")

			isSeparator := true
			for _, cell := range rawCells {
				cell = strings.TrimSpace(cell)

				// A valid separator column segment must only contain dash/colon/space characters
				// AND must contain at least one dash to confirm it's not an empty cell or pure spaces.
				if !separatorRe.MatchString(cell) || !strings.Contains(cell, "-") {
					isSeparator = false
					break
				}
			}

			if isSeparator {
				continue // Skip the separator row
			}

			// If we are here, it's a data row (header or content)
			rowVals := make([]string, 0, len(rawCells))
			for _, cell := range rawCells {
				rowVals = append(rowVals, strings.TrimSpace(cell))
			}

			if expectedCols == 0 && len(rowVals) > 0 {
				// The first valid row determines the column count
				expectedCols = len(rowVals)
			}

			// Ensure all data rows match the expected column count
			if expectedCols > 0 {
				if len(rowVals) < expectedCols {
					// Pad with empty strings
					for i := len(rowVals); i < expectedCols; i++ {
						rowVals = append(rowVals, "")
					}
				} else if len(rowVals) > expectedCols {
					// Truncate
					rowVals = rowVals[:expectedCols]
				}
			}

			rows = append(rows, rowVals)
		} else {
			// Once we hit a non-table line, we stop parsing the table
			if len(rows) > 0 {
				break
			}
		}
	}

	return rows, nil
}
