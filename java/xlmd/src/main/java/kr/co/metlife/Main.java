package kr.co.metlife;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

/**
 * XLMD interactive converter:
 * User pastes tab-delimited (copied-from-Excel) data into the terminal.
 * The app prints out the equivalent Markdown table.
 */
public class Main {

    private static final String BANNER = """
                     __  __     __         __    __     _____   \s
                    /\\_\\_\\_\\   /\\ \\       /\\ "-./  \\   /\\  __-. \s
                    \\/_/\\_\\/_  \\ \\ \\____  \\ \\ \\-./\\ \\  \\ \\ \\/\\ \\\s
                      /\\_\\/\\_\\  \\ \\_____\\  \\ \\_\\ \\ \\_\\  \\ \\____-\s
                      \\/_/\\/_/   \\/_____/   \\/_/  \\/_/   \\/____/\s
                     \s""";

    public static void main(String[] args) {
        System.out.println(BANNER);
        System.out.println("Paste the copied data from Excel below, then press ENTER twice when done:\n");

        List<String[]> rows = readUserPastedData();
        if (rows.isEmpty()) {
            System.out.println("No data received. Exiting.");
            return;
        }

        String markdown = convertToMarkdown(rows);
        System.out.println("\nConverted Markdown Table:\n");
        System.out.println(markdown);
    }

    /** Reads multiline user input (tab- or comma-delimited) until an empty line is entered. */
    private static List<String[]> readUserPastedData() {
        List<String[]> rows = new ArrayList<>();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(System.in))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.trim().isEmpty()) break; // blank line signals end of paste
                // Split on tab (preferred when pasting from Excel), fallback to comma
                String[] cells = line.split("\t", -1);
                if (cells.length == 1) cells = line.split(",", -1);
                rows.add(cells);
            }
        } catch (Exception e) {
            System.err.println("Error reading input: " + e.getMessage());
        }
        return rows;
    }

    /** Converts parsed rows into Markdown table format. */
    private static String convertToMarkdown(List<String[]> rows) {
        if (rows.isEmpty()) return "";

        int colCount = 0;
        for (String[] row : rows) {
            if (row.length > colCount) colCount = row.length;
        }

        // pad all rows to equal length
        for (int i = 0; i < rows.size(); i++) {
            String[] r = rows.get(i);
            if (r.length < colCount) {
                String[] padded = new String[colCount];
                System.arraycopy(r, 0, padded, 0, r.length);
                for (int j = r.length; j < colCount; j++) padded[j] = "";
                rows.set(i, padded);
            }
        }

        StringBuilder sb = new StringBuilder();
        // Header
        String[] header = rows.get(0);
        sb.append("| ").append(String.join(" | ", header)).append(" |\n");

        // Separator
        sb.append("|").append(" --- |".repeat(colCount)).append("\n");

        // Body
        for (int i = 1; i < rows.size(); i++) {
            sb.append("| ").append(String.join(" | ", rows.get(i))).append(" |\n");
        }

        return sb.toString();
    }
}
