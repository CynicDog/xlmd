package kr.co.metlife;

import kr.co.metlife.excel.ExcelReader;
import kr.co.metlife.excel.ExcelWriter;
import kr.co.metlife.excel.model.SheetData;
import kr.co.metlife.markdown.MarkdownReader;
import kr.co.metlife.markdown.MarkdownWriter;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.*;
import java.util.List;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit 5 integration tests for verifying Markdown <-> Excel conversions.
 *
 * Directory layout under src/test/resources/sample:
 *   iris/  → Excel (.xlsx) → Markdown (.md)
 *   sales/ → Markdown (.md) → Excel (.xlsx)
 */
public class ConvertTest {

    private static final Path RESOURCE_PATH = Paths.get("src", "test", "resources", "sample");

    /**
     * Test Excel → Markdown conversion.
     * Reads a real .xlsx, converts it, and compares to expected Markdown output.
     */
    @Test
    void testExcelToMarkdown(@TempDir Path tempDir) throws Exception {
        Path excelInput = RESOURCE_PATH.resolve("iris/in.xlsx");
        Path expectedMarkdown = RESOURCE_PATH.resolve("iris/out.md");
        Path actualMarkdown = tempDir.resolve("iris_out_actual.md");

        System.out.println("Testing Excel → Markdown: " + excelInput);

        // Read Excel file
        ExcelReader reader = new ExcelReader(excelInput.toString());
        List<SheetData> sheets = reader.read();
        assertFalse(sheets.isEmpty(), "ExcelReader should have extracted sheets.");

        // Write Markdown
        MarkdownWriter.writeMarkdown(actualMarkdown.toString(), sheets);
        assertTrue(Files.exists(actualMarkdown), "MarkdownWriter should create output file.");

        // Compare contents (normalized)
        String got = normalize(Files.readString(actualMarkdown));
        String want = normalize(Files.readString(expectedMarkdown));

        assertEquals(want, got, "Markdown output must match expected content.");
    }

    /**
     * Test Markdown → Excel conversion.
     * Reads a real Markdown table, writes Excel, ensures it’s valid, and re-reads it.
     */
    @Test
    void testMarkdownToExcel(@TempDir Path tempDir) throws Exception {
        Path markdownInput = RESOURCE_PATH.resolve("sales/in.md");
        Path excelExpected = RESOURCE_PATH.resolve("sales/out.xlsx");
        Path excelActual = tempDir.resolve("sales_out_actual.xlsx");

        System.out.println("Testing Markdown → Excel: " + markdownInput);

        // Read Markdown
        MarkdownReader mdReader = new MarkdownReader();
        List<SheetData> wantSheets = mdReader.readMarkdown(markdownInput.toString());
        assertFalse(wantSheets.isEmpty(), "MarkdownReader should extract at least one table.");

        // Write Excel file
        ExcelWriter writer = new ExcelWriter(excelActual.toString(), wantSheets);
        writer.write();
        assertTrue(Files.exists(excelActual), "ExcelWriter should produce an XLSX file.");

        // Verify the Excel file is a valid ZIP (XLSX container)
        try (ZipFile zip = new ZipFile(excelActual.toFile())) {
            assertNotNull(zip.getEntry("[Content_Types].xml"), "Missing [Content_Types].xml");
            assertNotNull(zip.getEntry("xl/workbook.xml"), "Missing workbook.xml");
            assertTrue(zip.stream().anyMatch(e -> e.getName().startsWith("xl/worksheets/sheet")),
                    "No worksheet files found in ZIP.");
        }

        // Read back the generated XLSX
        ExcelReader excelReader = new ExcelReader(excelActual.toString());
        List<SheetData> gotSheets = excelReader.read();
        assertFalse(gotSheets.isEmpty(), "Generated Excel file should be readable.");

        // Compare data content round-trip
        assertSheetDataEquals(wantSheets, gotSheets);
    }

    /** Normalize Markdown text to avoid whitespace differences. */
    private String normalize(String text) {
        return text.trim().replaceAll("\\s+", " ");
    }

    /** Deep equality check for SheetData contents. */
    private void assertSheetDataEquals(List<SheetData> expected, List<SheetData> actual) {
        assertEquals(expected.size(), actual.size(), "Sheet count mismatch.");
        for (int i = 0; i < expected.size(); i++) {
            SheetData want = expected.get(i);
            SheetData got = actual.get(i);
            assertEquals(want.getName(), got.getName(), "Sheet name mismatch at index " + i);
            assertEquals(want.getRows(), got.getRows(),
                    "Data mismatch in sheet: " + want.getName());
        }
    }
}
