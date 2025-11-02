package excel

import "encoding/xml"

// SheetData is the primary application model entity.
// It represents the simple, format-agnostic, two-dimensional table data
// extracted from or prepared for an Excel or Markdown file.
type SheetData struct {
	Name string
	Rows [][]string // Data content: row index -> column index -> cell value
}

// Worksheet represents the root element of an individual sheet's XML file (e.g., sheet1.xml).
// This structure is used for marshalling and unmarshalling the entire worksheet content.
type Worksheet struct {
	XMLName   xml.Name     `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetData SheetDataXML `xml:"sheetData"`
}

// SheetDataXML represents the mandatory <sheetData> container within the Worksheet XML.
// It holds all the row definitions for the sheet.
type SheetDataXML struct {
	Rows []Row `xml:"row"`
}

// Row represents a single <row> element in the XML, defined by its 1-based index (R).
type Row struct {
	R     int    `xml:"r,attr"` // Row index (1-based, required for XLSX structure)
	Cells []Cell `xml:"c"`
}

// Cell represents a single <c> element. This structure maps the technical XML attributes
// necessary for cell formatting and referencing (Ref, Type, Val).
type Cell struct {
	Ref  string `xml:"r,attr"`           // e.g., "A1", "B5" - Required cell reference
	Type string `xml:"t,attr,omitempty"` // "s" for shared string, otherwise numeric/empty
	Val  string `xml:"v"`                // The cell's value or the index (if type="s")
}

// WorkbookXML represents the root element of the workbook.xml file, detailing the overall workbook structure.
type WorkbookXML struct {
	XMLName xml.Name  `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	Sheets  SheetsXML `xml:"sheets"`
}

// SheetsXML represents the <sheets> container, listing all individual sheets in the workbook.
type SheetsXML struct {
	Sheet []SheetXMLInner `xml:"sheet"`
}

// SheetXMLInner defines the reference properties for a single sheet within workbook.xml.
type SheetXMLInner struct {
	Name    string `xml:"name,attr"`
	SheetID int    `xml:"sheetId,attr"`
	// RID is the relationship ID linking to the actual sheet XML file (e.g., rId1).
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// SST (Shared String Table) represents the root element of sharedStrings.xml.
// This table stores all unique strings/text used across the entire workbook.
type SST struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`       // Total number of strings
	UniqueCount int      `xml:"uniqueCount,attr"` // Number of unique strings
	SI          []SI     `xml:"si"`
}

// SI stands for Shared String Item. It represents the <si> tag, which contains the string text (<t>).
type SI struct {
	T string `xml:"t"` // The actual string value
}
