package kr.co.metlife;

import kr.co.metlife.excel.ExcelReader;
import kr.co.metlife.excel.ExcelWriter;
import kr.co.metlife.excel.model.SheetData;
import kr.co.metlife.markdown.MarkdownReader;
import kr.co.metlife.markdown.MarkdownWriter;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.*;
import java.util.List;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.*;

class ConvertTest {

    @TempDir
    Path tempDir;

    /** Excel → Markdown: simulate user paste input */
    @Test
    void testExcelToMarkdownModule() throws Exception {
        String simulatedInput = """
                sepal.length\tsepal.width\tpetal.length\tpetal.width\tvariety
                5.1\t3.5\t1.4\t0.2\tSetosa
                4.9\t3.0\t1.4\t0.2\tSetosa
                
                """; // blank lines terminate input

        InputStream originalIn = System.in;
        System.setIn(new ByteArrayInputStream(simulatedInput.getBytes()));

        List<SheetData> sheets = new ExcelReader().readFromUserInput();
        assertFalse(sheets.isEmpty(), "Should read at least one sheet");

        Path outputFile = tempDir.resolve("converted.md");
        MarkdownWriter writer = new MarkdownWriter();
        writer.writeMarkdown(outputFile.toString(), sheets);

        System.setIn(originalIn);

        assertTrue(Files.exists(outputFile), "Markdown output should exist");
        String content = Files.readString(outputFile);
        assertTrue(content.contains("| sepal.length |"), "Markdown header should exist");
    }

    /** Markdown → Excel: simulate user paste input */
    @Test
    void testMarkdownToExcelModule() throws Exception {
        String markdownInput = """
                | sepal.length | sepal.width | petal.length | petal.width | variety |
                | --- | --- | --- | --- | --- |
                | 5.1 | 3.5 | 1.4 | 0.2 | Setosa |
                | 4.9 | 3.0 | 1.4 | 0.2 | Setosa |
                
                """; // blank lines terminate input

        List<SheetData> sheets = new MarkdownReader().fromMarkdownString(markdownInput);
        assertFalse(sheets.isEmpty(), "Should parse at least one table");

        Path outputExcel = tempDir.resolve("converted.xlsx");
        ExcelWriter writer = new ExcelWriter(outputExcel.toString(), sheets);
        writer.write();

        assertTrue(Files.exists(outputExcel), "Excel output should exist");

        // Validate XLSX container
        try (ZipFile zip = new ZipFile(outputExcel.toFile())) {
            assertNotNull(zip.getEntry("[Content_Types].xml"), "Missing [Content_Types].xml");
            assertNotNull(zip.getEntry("xl/workbook.xml"), "Missing workbook.xml");
            assertTrue(zip.stream().anyMatch(e -> e.getName().startsWith("xl/worksheets/sheet")),
                    "No worksheet files found");
        }
    }
}
