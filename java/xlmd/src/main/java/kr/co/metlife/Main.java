package kr.co.metlife;

import kr.co.metlife.excel.ExcelReader;
import kr.co.metlife.excel.ExcelWriter;
import kr.co.metlife.excel.model.SheetData;
import kr.co.metlife.markdown.MarkdownReader;
import kr.co.metlife.markdown.MarkdownWriter;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.List;
import javax.xml.stream.XMLStreamException;

/**
 * Main command-line application entry point for XLMD.
 * Parses -i and -o flags, determines the conversion direction, and executes the logic.
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
        // Simple argument parsing for -i and -o
        String inputFile = null;
        String outputFile = null;

        for (int i = 0; i < args.length; i++) {
            if (args[i].equalsIgnoreCase("-i") && i + 1 < args.length) {
                inputFile = args[i + 1];
            } else if (args[i].equalsIgnoreCase("-o") && i + 1 < args.length) {
                outputFile = args[i + 1];
            }
        }

        if (inputFile == null || outputFile == null) {
            printUsage();
            System.exit(1);
        }

        runConversion(inputFile, outputFile);
    }

    private static void printUsage() {
        System.out.println(BANNER);
        System.out.println("Usage: java kr.co.metlife.cli.XLMD -i <input_file> -o <output_file>");
        System.out.println("\nExample: ");
        System.out.println("  Convert Excel to Markdown: XLMD -i data.xlsx -o output.md");
        System.out.println("  Convert Markdown to Excel: XLMD -i report.md -o table.xlsx");
        System.out.println("\nNote: Conversion direction is inferred from file extensions.");
    }

    private static void runConversion(String inputFile, String outputFile) {
        String inExt = getExtension(inputFile);
        String outExt = getExtension(outputFile);
        String direction = null;

        if (".xlsx".equalsIgnoreCase(inExt) && ".md".equalsIgnoreCase(outExt)) {
            direction = "excel2md";
        } else if (".md".equalsIgnoreCase(inExt) && ".xlsx".equalsIgnoreCase(outExt)) {
            direction = "md2excel";
        }

        if (direction == null) {
            System.err.println("\nError: Invalid file combination.");
            System.err.println("Must be either Excel → Markdown (.xlsx → .md) or Markdown → Excel (.md → .xlsx)");
            System.exit(1);
        }

        System.out.println(BANNER);
        System.out.printf("Convert %s input file to %s output file (%s)\n\n", inputFile, outputFile, direction);

        try {
            if ("excel2md".equals(direction)) {
                System.out.println("-> Reading Excel file...");
                ExcelReader reader = new ExcelReader(inputFile);
                List<SheetData> data = reader.read();
                System.out.println("-> Writing Markdown file...");
                MarkdownWriter.writeMarkdown(outputFile, data);
            } else { // md2excel
                System.out.println("-> Reading Markdown file...");
                MarkdownReader reader = new MarkdownReader();
                List<SheetData> data = reader.readMarkdown(inputFile);
                System.out.println("-> Writing Excel file...");
                ExcelWriter writer = new ExcelWriter(outputFile, data);
                writer.write();
            }
            System.out.println("\nSUCCESS: Conversion complete!");

        } catch (IOException e) {
            System.err.println("\nFATAL ERROR: A file system error occurred.");
            System.err.println("Details: " + e.getMessage());
            System.exit(1);
        } catch (XMLStreamException e) {
            System.err.println("\nFATAL ERROR: Failed to process XML content (XLSX/Markdown parsing error).");
            System.err.println("Details: " + e.getMessage());
            System.exit(1);
        } catch (Exception e) {
            System.err.println("\nFATAL ERROR: An unexpected error occurred.");
            e.printStackTrace(System.err);
            System.exit(1);
        }
    }

    private static String getExtension(String fileName) {
        String name = Paths.get(fileName).getFileName().toString();
        int lastDot = name.lastIndexOf('.');
        if (lastDot > 0) {
            return name.substring(lastDot).toLowerCase();
        }
        return "";
    }
}