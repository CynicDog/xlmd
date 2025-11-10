package kr.co.metlife.markdown;

import kr.co.metlife.excel.model.SheetData;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Reads a Markdown file, identifying sheets by "## Sheet Name" headers
 * and extracting subsequent Markdown tables into SheetData objects.
 */
public class MarkdownReader {

    private static final Pattern SHEET_HEADER_PATTERN = Pattern.compile("(?m)^##\\s*([^\\n]+)\\n");
    private static final Pattern SEPARATOR_ROW_PATTERN = Pattern.compile("^\\|(\\s*[-:]+\\s*\\|)+\\s*$");

    /**
     * Reads a Markdown file and converts its tables into a list of {@link SheetData} objects.
     * Each "## Sheet Name" header is treated as a separate sheet.
     *
     * @param filePath The path to the Markdown file to read.
     * @return A list of {@link SheetData} objects representing all parsed sheets.
     * @throws IOException If an error occurs while reading the file.
     */
    public List<SheetData> readMarkdown(String filePath) throws IOException {
        String content = new String(Files.readAllBytes(Paths.get(filePath)));
        List<SheetData> sheets = new ArrayList<>();

        Matcher matcher = SHEET_HEADER_PATTERN.matcher(content);
        List<String> sheetNames = new ArrayList<>();
        while (matcher.find()) {
            sheetNames.add(matcher.group(1).trim());
        }

        String[] contentParts = SHEET_HEADER_PATTERN.split(content, -1);

        if (sheetNames.isEmpty()) {
            List<List<String>> tableData = parseTable(content);
            if (!tableData.isEmpty()) {
                sheets.add(new SheetData("Sheet1", tableData));
            }
        } else {
            for (int i = 0; i < sheetNames.size(); i++) {
                if (i + 1 < contentParts.length) {
                    List<List<String>> tableData = parseTable(contentParts[i + 1]);
                    if (!tableData.isEmpty()) {
                        sheets.add(new SheetData(sheetNames.get(i), tableData));
                    }
                }
            }
        }
        return sheets;
    }

    /**
     * Parses a Markdown text block and extracts table rows into a list of string lists.
     * Each row corresponds to a table line, and each cell is trimmed of whitespace.
     *
     * @param textBlock The Markdown text block containing the table.
     * @return A list of rows, where each row is a list of cell values.
     */
    private List<List<String>> parseTable(String textBlock) {
        List<List<String>> tableRows = new ArrayList<>();
        int expectedCols = 0;
        String[] lines = textBlock.split("\\r?\\n");

        for (String line : lines) {
            line = line.trim();
            if (line.isEmpty() || !line.startsWith("|") || !line.endsWith("|")) {
                continue;
            }

            if (SEPARATOR_ROW_PATTERN.matcher(line).matches()) {
                continue;
            }

            String trimmed = line.substring(1, line.length() - 1).trim();
            List<String> rawCells = new ArrayList<>(Arrays.asList(trimmed.split("\\|", -1)));

            List<String> rowVals = new ArrayList<>();
            for (String cell : rawCells) {
                rowVals.add(cell.trim());
            }

            if (rowVals.isEmpty()) {
                continue;
            }

            if (expectedCols == 0) {
                expectedCols = rowVals.size();
            }

            if (rowVals.size() < expectedCols) {
                while (rowVals.size() < expectedCols) {
                    rowVals.add("");
                }
            } else if (rowVals.size() > expectedCols) {
                rowVals = rowVals.subList(0, expectedCols);
            }

            tableRows.add(rowVals);
        }
        return tableRows;
    }
}