package pkg_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"xlmd/pkg/excel"
	"xlmd/pkg/markdown"
)

func TestExcelToMarkdown(t *testing.T) {
	// Test description
	t.Log("Excel to Markdown conversion: iris.xlsx")

	input := filepath.Join("..", "sample", "iris", "in.xlsx")
	expected := filepath.Join("..", "sample", "iris", "out.md")

	// Read Excel in-memory
	sheets, err := excel.ReadExcel(input)
	if err != nil {
		t.Fatalf("failed to read Excel: %v", err)
	}

	// Convert to Markdown in-memory
	got := strings.TrimSpace(markdown.ToMarkdown(sheets))

	// Read expected Markdown
	wantBytes, err := os.ReadFile(expected)
	if err != nil {
		t.Fatalf("failed to read expected Markdown: %v", err)
	}
	want := strings.TrimSpace(string(wantBytes))

	// Assert equality
	if got != want {
		t.Errorf("Markdown output does not match expected.\n\nGot:\n%s\n\nWant:\n%s", got, want)
	}
}
