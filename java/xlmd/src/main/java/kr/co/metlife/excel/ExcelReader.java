package kr.co.metlife.excel;

import kr.co.metlife.excel.model.SheetData;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

/**
 * Reads pasted Excel tabular data (from clipboard paste in CLI)
 * and converts it into SheetData model.
 */
public class ExcelReader {

    /**
     * Reads multiline tab- or comma-delimited input from user (stdin)
     * until a blank line is entered.
     */
    public List<SheetData> readFromUserInput() {
        List<String[]> rows = new ArrayList<>();

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(System.in))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.trim().isEmpty()) break;

                // Split by tab (default from Excel copy), fallback to comma
                String[] cells = line.split("\t", -1);
                if (cells.length == 1) cells = line.split(",", -1);

                rows.add(cells);
            }
        } catch (Exception e) {
            System.err.println("Error reading input: " + e.getMessage());
        }

        List<SheetData> sheets = new ArrayList<>();
        if (!rows.isEmpty()) {
            sheets.add(new SheetData("Converted", rows));
        }
        return sheets;
    }
}
