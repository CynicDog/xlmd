package excel

// SheetData represents a sheet in Excel or Markdown
type SheetData struct {
	Name string
	Rows [][]string
}

// ReadExcel reads an Excel file and returns all sheets as SheetData
func ReadExcel(filePath string) ([]SheetData, error) {
	// TODO; Implementation goes here
	return nil, nil
}
