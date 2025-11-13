package kr.co.metlife;

import kr.co.metlife.excel.ExcelReader;
import kr.co.metlife.excel.ExcelWriter;
import kr.co.metlife.excel.model.SheetData;
import kr.co.metlife.markdown.MarkdownReader;
import kr.co.metlife.markdown.MarkdownWriter;

import java.util.List;
import java.util.Scanner;

/**
 * Interactive CLI entry point for XLMD
 * Two modes:
 *  1) Excel → Markdown  : Paste Excel cells → outputs converted.md
 *  2) Markdown → Excel  : Paste Markdown text → outputs converted.xlsx
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
        System.out.println("Choose conversion mode:");
        System.out.println("1) Excel → Markdown");
        System.out.println("2) Markdown → Excel");
        System.out.print("Enter choice (1 or 2): ");

        int choice = readChoice();

        switch (choice) {
            case 1 -> runExcelToMarkdown();
            case 2 -> runMarkdownToExcel(); // implemented later
            default -> System.out.println("Invalid choice. Exiting.");
        }
    }

    /** Reads user’s numeric menu choice */
    private static int readChoice() {
        Scanner sc = new Scanner(System.in);
        if (sc.hasNextInt()) {
            return sc.nextInt();
        }
        return -1;
    }

    /** === MODE 1: Excel → Markdown === */
    private static void runExcelToMarkdown() {
        System.out.println("\nPaste the copied data from Excel below, then press ENTER twice when done:\n");

        ExcelReader reader = new ExcelReader();
        List<SheetData> sheets = reader.readFromUserInput();

        if (sheets.isEmpty() || sheets.get(0).getRows().isEmpty()) {
            System.out.println("No data received. Exiting.");
            return;
        }

        MarkdownWriter writer = new MarkdownWriter();
        writer.writeMarkdown("converted.md", sheets);

        System.out.println("\n Conversion complete!");
        System.out.println("Markdown saved to: converted.md");
    }

    /** === MODE 2: Markdown → Excel === */
    private static void runMarkdownToExcel() {
        System.out.println("\nPaste your Markdown table(s) below, then press ENTER when done:\n");

        // Read multi-line input from user until two consecutive blank lines
        Scanner sc = new Scanner(System.in);
        StringBuilder inputBuilder = new StringBuilder();

        while (true) {
            String line = sc.nextLine();
            if (line.trim().isEmpty()) break; // stop at first blank line
            inputBuilder.append(line).append("\n");
        }

        String markdownInput = inputBuilder.toString().trim();

        if (markdownInput.isEmpty()) {
            System.out.println("No Markdown input detected. Exiting.");
            return;
        }

        // Parse Markdown into SheetData
        MarkdownReader mdReader = new MarkdownReader();
        List<SheetData> sheets = mdReader.fromMarkdownString(markdownInput);

        if (sheets.isEmpty()) {
            System.out.println("No table data found in Markdown. Exiting.");
            return;
        }

        // Write Excel file
        ExcelWriter writer = new ExcelWriter("converted.xlsx", sheets);
        try {
            writer.write();
            System.out.println("\nConversion complete!");
            System.out.println("Excel saved to: converted.xlsx");
        } catch (Exception e) {
            System.out.println("Failed to write Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

}
