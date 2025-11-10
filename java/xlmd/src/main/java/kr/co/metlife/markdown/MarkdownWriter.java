package kr.co.metlife.markdown;

import kr.co.metlife.excel.model.SheetData;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Converts a list of SheetData objects into a single Markdown string.
 * Each sheet is converted into a level 2 header (##) followed by a
 * GitHub Flavored Markdown (GFM) table.
 */
public class MarkdownWriter {

    /**
     * Writes the output Markdown content to the specified file path.
     * @param filePath The path where the Markdown file should be saved.
     * @param sheets The data structure containing all sheet data.
     * @throws IOException if writing the file fails.
     */
    public static void writeMarkdown(String filePath, List<SheetData> sheets) throws IOException {
        String markdownContent = toMarkdown(sheets);
        Files.write(Paths.get(filePath), markdownContent.getBytes());
    }

    /**
     * Converts the structured SheetData into a single Markdown string.
     * @param sheets The list of sheets to convert.
     * @return A single string containing all sheets formatted as Markdown.
     */
    public static String toMarkdown(List<SheetData> sheets) {
        StringBuilder sb = new StringBuilder();

        for (SheetData sheet : sheets) {
            // Skip sheets with no data
            if (sheet.getRows() == null || sheet.getRows().isEmpty()) {
                continue;
            }

            // Sheet Header: ## Sheet Name
            sb.append("## ").append(sheet.getName()).append("\n\n");

            List<List<String>> rows = sheet.getRows();

            // Determine column count (based on the longest row)
            int colCount = rows.stream()
                    .mapToInt(List::size)
                    .max()
                    .orElse(0);

            // If there are no columns, skip the table structure
            if (colCount == 0) {
                sb.append("\n");
                continue;
            }

            // Pad all rows to the determined column count
            List<List<String>> paddedRows = rows.stream()
                    .map(row -> padRow(row, colCount))
                    .collect(Collectors.toList());


            // Write Header Row (First row of data is used as the table header)
            List<String> header = paddedRows.get(0);
            sb.append("| ").append(String.join(" | ", header)).append(" |\n");

            // Write Separator Row (| --- | --- | ...)
            // We create a list of " --- " strings and join them with '|'.
            String separator = Collections.nCopies(colCount, "---").stream()
                    .collect(Collectors.joining(" | ", "| ", " |\n"));
            sb.append(separator);

            // Write Data Rows (The rest of the rows)
            // Start from index 1 as index 0 was used for the header.
            for (int i = 1; i < paddedRows.size(); i++) {
                List<String> row = paddedRows.get(i);
                // Wrap each cell value with spaces and join them with ' | '
                sb.append("| ").append(String.join(" | ", row)).append(" |\n");
            }

            // Add an extra newline for visual separation between tables
            sb.append("\n");
        }

        return sb.toString();
    }

    /**
     * Pads a given row with empty strings to ensure it reaches the specified number of columns.
     *
     * @param row The original list of cell values for a single row.
     * @param colCount The desired number of columns the row should have.
     * @return A new list representing the padded row, with empty strings ("") added
     *         if the original row had fewer cells than {@code colCount}.
     */
    private static List<String> padRow(List<String> row, int colCount) {
        if (row.size() >= colCount) {
            // If the row is already wide enough (or too wide), return a truncated copy
            return row.subList(0, colCount);
        }

        List<String> padded = new java.util.ArrayList<>(row);
        while (padded.size() < colCount) {
            padded.add(""); // Pad with empty strings
        }
        return padded;
    }
}
