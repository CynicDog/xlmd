package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// TODO: move shared Excel and XML struct definitions into a dedicated file (e.g. excel_types.go)
// These types are reused across both reading and writing logic and should be isolated
// to simplify maintenance and reduce duplication.

type SheetData struct {
	Name string
	Rows [][]string
}

type SheetXML struct {
	Rows []Row `xml:"row"`
}

type Row struct {
	R     int    `xml:"r,attr"` // Row index (1-based)
	Cells []Cell `xml:"c"`
}

type Cell struct {
	Ref  string `xml:"r,attr"`
	Type string `xml:"t,attr,omitempty"`
	Val  string `xml:"v"`
}

// Worksheet Main XML structure for a single sheet (e.g., sheet1.xml)
type Worksheet struct {
	XMLName   xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetData SheetXML `xml:"sheetData"`
}

// WorkbookXML lists all sheets and their relationships.
type WorkbookXML struct {
	XMLName xml.Name  `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	Sheets  SheetsXML `xml:"sheets"`
}

// SST (Shared String Table): Unique strings used in the workbook.
type SST struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	SI          []SI     `xml:"si"`
}

type SI struct {
	T string `xml:"t"` // Text
}

type SheetsXML struct {
	Sheet []SheetXMLInner `xml:"sheet"`
}

type SheetXMLInner struct {
	Name    string `xml:"name,attr"`
	SheetID int    `xml:"sheetId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// ReadExcel opens and reads a .xlsx spreadsheet file, returning its sheets and Cell data
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
			var sx SheetXML
			if err := xml.Unmarshal(xmlData, &sx); err != nil {
				return nil, fmt.Errorf("failed to parse sheet XML: %w", err)
			}

			var sheet SheetData
			sheet.Name = guessSheetName(f.Name)

			for _, r := range sx.Rows {
				var rowVals []string
				for _, c := range r.Cells {
					v := c.Val
					// When a Cell’s type attribute is “s”, its value represents an index
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
