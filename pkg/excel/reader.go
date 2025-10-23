package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

type SheetData struct {
	Name string
	Rows [][]string
}

type cell struct {
	Ref  string `xml:"r,attr"`
	Type string `xml:"t,attr"`
	Val  string `xml:"v"`
}

type row struct {
	Cells []cell `xml:"c"`
}

type sheetXML struct {
	Rows []row `xml:"sheetData>row"`
}

// ReadExcel opens and reads a .xlsx spreadsheet file, returning its sheets and cell data
// as a slice of SheetData.
func ReadExcel(filePath string) ([]SheetData, error) {
	zr, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open the xlsx file: %w", err)
	}
	defer zr.Close()

	var sharedStrings []string
	var sheets []SheetData

	// Locate and parse the sharedStrings.xml file if present.
	// This file contains all unique text strings used throughout the workbook.
	for _, f := range zr.File {
		if f.Name == "xl/sharedStrings.xml" {
			rc, _ := f.Open()
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			type sst struct {
				SI []struct {
					T string `xml:"t"`
				} `xml:"si"`
			}
			var s sst
			xml.Unmarshal(data, &s)
			for _, v := range s.SI {
				sharedStrings = append(sharedStrings, v.T)
			}
			break
		}
	}
	// Each worksheet is stored as an XML file under xl/worksheets/, typically named sheet1.xml,
	// sheet2.xml, and so on. The following loop extracts each of these sheets and converts its
	// XML representation into a SheetData structure. Cells that reference shared strings are
	// replaced with their resolved text values.
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/sheet") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			defer rc.Close()

			xmlData, _ := io.ReadAll(rc)
			var sx sheetXML
			if err := xml.Unmarshal(xmlData, &sx); err != nil {
				return nil, fmt.Errorf("failed to parse sheet XML: %w", err)
			}

			var sheet SheetData
			sheet.Name = guessSheetName(f.Name)

			for _, r := range sx.Rows {
				var rowVals []string
				for _, c := range r.Cells {
					v := c.Val
					// When a cell’s type attribute is “s”, its value represents an index
					// into the shared strings table. In that case, we replace the numeric
					// index with the corresponding string value.
					if c.Type == "s" {
						idx, _ := strconv.Atoi(v)
						if idx < len(sharedStrings) {
							v = sharedStrings[idx]
						}
					}
					rowVals = append(rowVals, v)
				}
				sheet.Rows = append(sheet.Rows, rowVals)
			}

			sheets = append(sheets, sheet)
		}
	}

	return sheets, nil
}

// guessSheetName extracts a simple worksheet name from its internal
// path inside the XLSX archive, removing the directory and ".xml"
// extension (e.g., "xl/worksheets/sheet1.xml" → "Sheet1").
func guessSheetName(path string) string {
	name := strings.TrimPrefix(path, "xl/worksheets/")
	name = strings.TrimSuffix(name, ".xml")
	return strings.Title(name)
}
