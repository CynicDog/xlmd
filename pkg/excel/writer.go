package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"strconv"
)

// toColName converts a 0-based column index (0, 1, 2...) into an Excel column letter ("A", "B", "C"...).
// For example, 0 -> "A", 25 -> "Z", 26 -> "AA".
func toColName(col int) string {
	if col < 0 {
		return ""
	}
	var name string
	for col >= 0 {
		name = string('A'+col%26) + name
		col = col/26 - 1
	}
	return name
}

// WriteExcel generates a complete XLSX file from the provided SheetData using standard Go libraries.
// It handles the creation of the ZIP archive structure and the necessary OpenXML parts.
func WriteExcel(filePath string, sheets []SheetData) error {
	sharedStrings := make([]string, 0)
	stringIndexMap := make(map[string]int)

	// The SST is required by the OpenXML specification to store all unique textual cell values.
	for _, sheet := range sheets {
		for _, row := range sheet.Rows {
			for _, cellValue := range row {
				// Only process non-empty strings
				if cellValue != "" {
					if _, exists := stringIndexMap[cellValue]; !exists {
						stringIndexMap[cellValue] = len(sharedStrings)
						sharedStrings = append(sharedStrings, cellValue)
					}
				}
			}
		}
	}

	// Create the output ZIP file and initialize the ZIP writer.
	zipFile, err := os.Create(filePath)
	if err != nil {
		return fmt.Errorf("failed to create zip file: %w", err)
	}
	defer zipFile.Close()

	zw := zip.NewWriter(zipFile)
	defer zw.Close()

	// writeXML is a helper function to marshal a Go struct into its XML representation
	// and write it as an entry within the XLSX ZIP archive. It ensures proper XML header
	// and indentation.
	writeXML := func(filename string, data interface{}) error {
		f, err := zw.Create(filename)
		if err != nil {
			return err
		}

		// Write XML header
		f.Write([]byte(xml.Header))

		enc := xml.NewEncoder(f)
		enc.Indent("", "  ")
		if err := enc.Encode(data); err != nil {
			return fmt.Errorf("failed to encode XML for %s: %w", filename, err)
		}
		return nil
	}

	// [Content_Types].xml (Defines MIME types for all parts)
	// This file defines the MIME type for every file part within the package (e.g., workbook, sheets, relationships).
	contentTypeXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
`
	for i := range sheets {
		contentTypeXML += fmt.Sprintf(`    <Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
`, i+1)
	}
	contentTypeXML += `</Types>`
	if f, err := zw.Create("[Content_Types].xml"); err != nil {
		return err
	} else {
		f.Write([]byte(contentTypeXML))
	}

	// _rels/.rels (Package Relationships)
	// This is the package-level relationship file, linking the package to the main workbook part.
	relsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
	if f, err := zw.Create("_rels/.rels"); err != nil {
		return err
	} else {
		f.Write([]byte(relsXML))
	}

	// xl/styles.xml (Required empty styles file)
	// An empty, minimal styles file is mandatory for Excel to open the file correctly.
	stylesXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>`
	if f, err := zw.Create("xl/styles.xml"); err != nil {
		return err
	} else {
		f.Write([]byte(stylesXML))
	}

	// xl/sharedStrings.xml
	// Writes the collected unique string data (SST) to the archive.
	sstData := SST{
		Count:       len(sharedStrings),
		UniqueCount: len(sharedStrings),
		SI:          make([]SI, len(sharedStrings)),
	}
	for i, s := range sharedStrings {
		sstData.SI[i] = SI{T: s}
	}

	if err := writeXML("xl/sharedStrings.xml", sstData); err != nil {
		return err
	}

	// xl/_rels/workbook.xml.rels (Workbook Relationships)
	// Defines the relationships for the workbook, linking it to all its component parts:
	// sheets, styles, and the shared string table.
	wbRelsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
`
	for i := range sheets {
		wbRelsXML += fmt.Sprintf(` <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>
`, i+1, i+1)
	}
	// rIdX+1 for styles, rIdX+2 for shared strings
	wbRelsXML += fmt.Sprintf(` <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
`, len(sheets)+1)
	wbRelsXML += fmt.Sprintf(` <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`, len(sheets)+2)

	if f, err := zw.Create("xl/_rels/workbook.xml.rels"); err != nil {
		return err
	} else {
		f.Write([]byte(wbRelsXML))
	}

	// xl/workbook.xml
	wbData := WorkbookXML{
		Sheets: SheetsXML{
			Sheet: make([]SheetXMLInner, len(sheets)),
		},
	}
	for i, sheet := range sheets {
		sheetName := sheet.Name
		if sheetName == "" {
			sheetName = fmt.Sprintf("Sheet%d", i+1)
		}
		wbData.Sheets.Sheet[i] = SheetXMLInner{
			Name:    sheetName,
			SheetID: i + 1,
			RID:     fmt.Sprintf("rId%d", i+1),
		}
	}
	if err := writeXML("xl/workbook.xml", wbData); err != nil {
		return err
	}

	// xl/worksheets/sheetX.xml
	// Generate XML for each individual worksheet (xl/worksheets/sheetX.xml).
	for i, sheet := range sheets {
		xmlRows := make([]Row, 0, len(sheet.Rows))

		for rIdx, row := range sheet.Rows {
			rowNum := rIdx + 1 // 1-based index
			xmlCells := make([]Cell, 0, len(row))

			// Find the column count for this Row (max index)
			maxColIndex := -1
			for j, val := range row {
				if val != "" {
					maxColIndex = j
				}
			}

			// Only write cells up to the last non-empty column
			for cIdx := 0; cIdx <= maxColIndex; cIdx++ {
				cellValue := row[cIdx]
				if cellValue == "" {
					// Excel omits empty cells in the XML
					continue
				}

				colName := toColName(cIdx)
				cellRef := fmt.Sprintf("%s%d", colName, rowNum)

				// All strings are stored as shared strings
				stringIndex := stringIndexMap[cellValue]

				xmlCells = append(xmlCells, Cell{
					Ref:  cellRef,
					Type: "s",
					Val:  strconv.Itoa(stringIndex),
				})
			}

			// Only include rows that have at least one Cell
			if len(xmlCells) > 0 {
				xmlRows = append(xmlRows, Row{
					R:     rowNum,
					Cells: xmlCells,
				})
			}
		}

		// Final assembly of the worksheet XML.
		wsData := Worksheet{
			SheetData: SheetDataXML{
				Rows: xmlRows,
			},
		}

		filename := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)
		if err := writeXML(filename, wsData); err != nil {
			return err
		}
	}

	return nil
}
