package pkg_test

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"xlmd/pkg/excel"
	"xlmd/pkg/markdown"
)

// TestExcelToMarkdown verifies the end-to-end fidelity of the Excel (XLSX) to Markdown conversion.
func TestExcelToMarkdown(t *testing.T) {
	t.Log("Excel to Markdown conversion: iris.xlsx")

	input := filepath.Join("..", "sample", "iris", "in.xlsx")
	expected := filepath.Join("..", "sample", "iris", "out.md")

	sheets, err := excel.ReadExcel(input)
	if err != nil {
		t.Fatalf("failed to read Excel: %v", err)
	}

	got := strings.TrimSpace(markdown.ToMarkdown(sheets))

	wantBytes, err := os.ReadFile(expected)
	if err != nil {
		t.Fatalf("failed to read expected Markdown: %v", err)
	}
	want := strings.TrimSpace(string(wantBytes))

	if got != want {
		t.Errorf("Markdown output does not match expected.\n\nGot:\n%s\n\nWant:\n%s", got, want)
	}
}

// TestMarkdownToExcel tests the conversion from a Markdown file back into an Excel (XLSX) file.
// It verifies the fidelity of the conversion by comparing the generated sheet data with the expected sheet data.
func TestMarkdownToExcel(t *testing.T) {
	t.Log("Testing Markdown to Excel conversion: iris.md")

	input := filepath.Join("..", "sample", "sales", "in.md")
	expected := filepath.Join("..", "sample", "sales", "out.xlsx")
	tempOutput := filepath.Join(os.TempDir(), "test_md_to_excel_output.xlsx")

	// Ensure the temporary output file is cleaned up after the test finishes.
	defer os.Remove(tempOutput)

	sheets, err := markdown.ReadMarkdown(input)
	if err != nil {
		t.Fatalf("markdown.ReadMarkdown failed: %v", err)
	}
	if err := excel.WriteExcel(tempOutput, sheets); err != nil {
		t.Fatalf("excel.WriteExcel failed: %v", err)
	}

	wantSheets, err := excel.ReadExcel(expected)
	if err != nil {
		t.Fatalf("failed to read expected Excel file: %v", err)
	}
	gotSheets, err := excel.ReadExcel(tempOutput)
	if err != nil {
		t.Fatalf("failed to read generated Excel file: %v", err)
	}

	if len(gotSheets) != len(wantSheets) {
		t.Fatalf("Sheet count mismatch. Got %d sheets, Want %d sheets.", len(gotSheets), len(wantSheets))
	}

	for i := range gotSheets {
		gotSheet := gotSheets[i]
		wantSheet := wantSheets[i]

		if gotSheet.Name != wantSheet.Name {
			t.Errorf("Sheet %d name mismatch. Got %q, Want %q", i+1, gotSheet.Name, wantSheet.Name)
			continue
		}

		// Use DeepEqual to verify the content (Rows) of the sheets match exactly.
		if !reflect.DeepEqual(gotSheet.Rows, wantSheet.Rows) {
			t.Errorf("Sheet %q content mismatch.", gotSheet.Name)
			t.Logf("Got data mismatch on sheet %q", gotSheet.Name)
		}
	}
}
