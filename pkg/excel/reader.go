package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"strconv"
	"strings"
)

// colRefToIndex converts an Excel column reference (e.g., "A1", "C5") to a 0-based index (A=0).
func colRefToIndex(ref string) int {
	colStr := ""
	for _, char := range ref {
		if char >= 'A' && char <= 'Z' {
			colStr += string(char)
		} else {
			break
		}
	}

	index := 0
	for i, char := range colStr {
		power := len(colStr) - i - 1
		index += (int(char) - 'A' + 1) * int(math.Pow(26, float64(power)))
	}
	// Convert 1-based index (1=A) to 0-based index (0=A)
	return index - 1
}

// ReadExcel opens and reads a .xlsx spreadsheet file, returning its sheets and Cell data.
func ReadExcel(filePath string) ([]SheetData, error) {
	zr, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open the xlsx file: %w", err)
	}
	defer zr.Close()

	var sharedStrings []string
	var sheets []SheetData

	// Locate and parse sharedStrings.xml
	for _, f := range zr.File {
		if f.Name == "xl/sharedStrings.xml" {
			rc, openErr := f.Open()
			if openErr != nil {
				return nil, fmt.Errorf("failed to open sharedStrings.xml: %w", openErr)
			}
			defer rc.Close()
			data, readErr := io.ReadAll(rc)
			if readErr != nil {
				return nil, fmt.Errorf("failed to read sharedStrings.xml: %w", readErr)
			}

			var sst SST
			if xml.Unmarshal(data, &sst) == nil {
				for _, si := range sst.SI {
					sharedStrings = append(sharedStrings, si.T)
				}
			}
			break
		}
	}

	// Iterate through worksheet XML files and extract data.
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/sheet") && strings.HasSuffix(f.Name, ".xml") {
			rc, openErr := f.Open()
			if openErr != nil {
				return nil, fmt.Errorf("failed to open sheet XML %s: %w", f.Name, openErr)
			}
			defer rc.Close()

			xmlData, readErr := io.ReadAll(rc)
			if readErr != nil {
				return nil, fmt.Errorf("failed to read sheet XML %s: %w", f.Name, readErr)
			}

			var ws Worksheet
			if err := xml.Unmarshal(xmlData, &ws); err != nil {
				return nil, fmt.Errorf("failed to parse sheet XML %s: %w", f.Name, err)
			}

			var sheet SheetData
			sheet.Name = guessSheetName(f.Name)

			for _, r := range ws.SheetData.Rows {
				if len(r.Cells) == 0 {
					continue
				}

				// Find the max column index required for this row based on cell references
				maxColIndex := 0
				for _, c := range r.Cells {
					colIndex := colRefToIndex(c.Ref)
					if colIndex > maxColIndex {
						maxColIndex = colIndex
					}
				}

				rowVals := make([]string, maxColIndex+1)

				// Populate the row slice by placing values at the correct column index.
				for _, c := range r.Cells {
					colIndex := colRefToIndex(c.Ref)

					v := c.Val
					// Resolve shared strings if type="s"
					if c.Type == "s" {
						idx, atoiErr := strconv.Atoi(v)
						if atoiErr == nil && idx >= 0 && idx < len(sharedStrings) {
							v = sharedStrings[idx]
						} else {
							v = ""
						}
					}

					rowVals[colIndex] = v
				}

				sheet.Rows = append(sheet.Rows, rowVals)
			}

			sheets = append(sheets, sheet)
		}
	}

	return sheets, nil
}

// guessSheetName extracts a simple worksheet name from its internal path.
func guessSheetName(path string) string {
	name := strings.TrimPrefix(path, "xl/worksheets/")
	name = strings.TrimSuffix(name, ".xml")
	return strings.Title(name)
}
