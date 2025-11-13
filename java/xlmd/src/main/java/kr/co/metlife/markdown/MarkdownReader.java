package kr.co.metlife.markdown;

import kr.co.metlife.excel.model.SheetData;

import java.util.*;
import java.util.regex.*;

/**
 * Reads Markdown table string (from CLI input) and converts to SheetData
 */
public class MarkdownReader {

    private static final Pattern SEPARATOR_ROW_PATTERN = Pattern.compile("^\\|(\\s*[-:]+\\s*\\|)+\\s*$");

    /**
     * Reads a Markdown string and returns one sheet (Sheet1) with the table
     */
    public List<SheetData> fromMarkdownString(String markdown) {
        List<List<String>> tableRows = parseTable(markdown);
        List<String[]> rowsArray = new ArrayList<>();
        for (List<String> row : tableRows) {
            rowsArray.add(row.toArray(new String[0]));
        }
        SheetData sheet = new SheetData("Sheet1", rowsArray);
        return List.of(sheet);
    }

    private List<List<String>> parseTable(String text) {
        List<List<String>> tableRows = new ArrayList<>();
        String[] lines = text.split("\\r?\\n");
        int expectedCols = 0;

        for (String line : lines) {
            line = line.trim();
            if (line.isEmpty() || !line.startsWith("|") || !line.endsWith("|")) continue;
            if (SEPARATOR_ROW_PATTERN.matcher(line).matches()) continue;

            String trimmed = line.substring(1, line.length() - 1);
            String[] rawCells = trimmed.split("\\|", -1);
            List<String> rowVals = new ArrayList<>();
            for (String cell : rawCells) rowVals.add(cell.trim());

            if (rowVals.isEmpty()) continue;
            if (expectedCols == 0) expectedCols = rowVals.size();
            while (rowVals.size() < expectedCols) rowVals.add("");
            if (rowVals.size() > expectedCols) rowVals = rowVals.subList(0, expectedCols);

            tableRows.add(rowVals);
        }

        return tableRows;
    }
}
