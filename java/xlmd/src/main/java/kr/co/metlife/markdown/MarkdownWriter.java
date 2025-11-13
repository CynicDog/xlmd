package kr.co.metlife.markdown;

import kr.co.metlife.excel.model.SheetData;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Converts SheetData into Markdown table syntax and writes to file.
 */
public class MarkdownWriter {

    /**
     * Writes all provided sheets as Markdown tables into a .md file.
     */
    public void writeMarkdown(String filePath, List<SheetData> sheets) {
        StringBuilder sb = new StringBuilder();

        for (SheetData sheet : sheets) {
            if (sheet.getRows().isEmpty()) continue;

            sb.append("## ").append(sheet.getName()).append("\n\n");

            int colCount = sheet.getMaxColumnCount();

            // Header
            String[] header = sheet.getRows().get(0);
            sb.append("| ").append(String.join(" | ", padRow(header, colCount))).append(" |\n");

            // Separator
            sb.append("|").append(" --- |".repeat(colCount)).append("\n");

            // Body
            for (int i = 1; i < sheet.getRows().size(); i++) {
                String[] row = padRow(sheet.getRows().get(i), colCount);
                sb.append("| ").append(String.join(" | ", row)).append(" |\n");
            }

            sb.append("\n");
        }

        try {
            Files.writeString(Path.of(filePath), sb.toString());
        } catch (IOException e) {
            System.err.println("Error writing Markdown file: " + e.getMessage());
        }
    }

    /** Pads a row to match the column count with empty cells */
    private String[] padRow(String[] row, int colCount) {
        if (row.length == colCount) return row;
        String[] padded = new String[colCount];
        System.arraycopy(row, 0, padded, 0, row.length);
        for (int i = row.length; i < colCount; i++) padded[i] = "";
        return padded;
    }
}
