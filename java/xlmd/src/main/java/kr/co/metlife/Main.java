package kr.co.metlife;

import kr.co.metlife.excel.ExcelReader;
import kr.co.metlife.excel.ExcelWriter;
import kr.co.metlife.excel.model.SheetData;
import kr.co.metlife.markdown.MarkdownReader;
import kr.co.metlife.markdown.MarkdownWriter;

import java.util.List;
import java.util.Scanner;

/**
 * Interactive CLI의 entry point입니다.
 * 두 가지 모드를 지원합니다:
 *  1) Excel → Markdown  : Excel 셀을 붙여넣으면 converted.md로 출력
 *  2) Markdown → Excel  : Markdown 텍스트를 붙여넣으면 converted.xlsx로 출력
 */
public class Main {

    private static final String BANNER = """
                     __  __     __         __    __     _____   \s
                    /\\_\\_\\_\\   /\\ \\       /\\ "-./  \\   /\\  __-. \s
                    \\/_/\\_\\/_  \\ \\ \\____  \\ \\ \\-./\\ \\  \\ \\ \\/\\ \\\s
                      /\\_\\/\\_\\  \\ \\_____\\  \\ \\_\\ \\ \\_\\  \\ \\____-\s
                      \\/_/\\/_/   \\/_____/   \\/_/  \\/_/   \\/____/\s
                     \s""";

    /**
     * 프로그램의 entry point로서, 사용자에게 변환 모드를 선택하도록 안내하고
     * 선택에 따라 Excel → Markdown 또는 Markdown → Excel 변환을 실행합니다.
     *
     * @param args 커맨드라인 인수 (미사용)
     */
    public static void main(String[] args) {
        System.out.println(BANNER);
        System.out.println("변환 모드를 선택하세요:");
        System.out.println("1) Excel → Markdown");
        System.out.println("2) Markdown → Excel");
        System.out.print("입력: ");

        int choice = readChoice();

        switch (choice) {
            case 1 -> runExcelToMarkdown();
            case 2 -> runMarkdownToExcel(); // implemented later
            default -> System.out.println("잘못된 선택입니다. 프로그램을 종료합니다.");
        }
    }

    /**
     * 사용자가 입력한 메뉴 선택 숫자를 읽습니다.
     *
     * @return 사용자가 입력한 숫자 선택, 유효하지 않으면 -1
     */
    private static int readChoice() {
        Scanner sc = new Scanner(System.in);
        if (sc.hasNextInt()) {
            return sc.nextInt();
        }
        return -1;
    }

    /**
     * 모드 1: Excel → Markdown 변환을 수행합니다.
     * 사용자가 Excel에서 복사한 데이터를 붙여넣으면
     * converted.md 파일로 Markdown 형식으로 저장됩니다.
     */
    private static void runExcelToMarkdown() {
        System.out.println("\nExcel에서 복사한 데이터를 붙여넣고 ENTER를 두 번 누르세요:\n");

        ExcelReader reader = new ExcelReader();
        List<SheetData> sheets = reader.readFromUserInput();

        if (sheets.isEmpty() || sheets.get(0).getRows().isEmpty()) {
            System.out.println("데이터가 입력되지 않았습니다. 프로그램을 종료합니다.");
            return;
        }

        MarkdownWriter writer = new MarkdownWriter();
        writer.writeMarkdown("converted.md", sheets);

        System.out.println("Markdown 파일이 저장되었습니다: converted.md");
    }

    /**
     * 모드 2: Markdown → Excel 변환을 수행합니다.
     * 사용자가 Markdown 테이블을 붙여넣으면
     * converted.xlsx 파일로 Excel 형식으로 저장됩니다.
     */
    private static void runMarkdownToExcel() {
        System.out.println("\nMarkdown 테이블을 붙여넣고 완료되면 ENTER를 누르세요:\n");

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
            System.out.println("Markdown 테이블이 입력되지 않았습니다. 프로그램을 종료합니다.");
            return;
        }

        // Parse Markdown into SheetData
        MarkdownReader mdReader = new MarkdownReader();
        List<SheetData> sheets = mdReader.fromMarkdownString(markdownInput);

        if (sheets.isEmpty()) {
            System.out.println("Markdown에서 테이블 데이터를 찾을 수 없습니다. 프로그램을 종료합니다.");
            return;
        }

        // Write Excel file
        ExcelWriter writer = new ExcelWriter("converted.xlsx", sheets);
        try {
            writer.write();
            System.out.println("Excel 파일이 저장되었습니다: converted.xlsx");
        } catch (Exception e) {
            System.out.println("Excel 파일 작성에 실패했습니다: " + e.getMessage());
            e.printStackTrace();
        }
    }
}