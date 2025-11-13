package kr.co.metlife.markdown;

import kr.co.metlife.excel.model.SheetData;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * SheetData를 Markdown 테이블 문법으로 변환하여 파일에 작성합니다.
 */
public class MarkdownWriter {

    /**
     * 전달받은 모든 시트(SheetData) 정보를 Markdown 테이블 형식으로 변환하여
     * .md 파일로 작성합니다.
     *
     * 첫 번째 행을 헤더로 사용하고, 부족한 컬럼은 빈 문자열로 패딩됩니다.
     *
     * @param filePath 출력할 Markdown 파일 경로
     * @param sheets 변환할 SheetData 리스트
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

    /**
     * 행(row)의 길이가 컬럼 수(colCount)에 맞지 않을 경우,
     * 부족한 부분을 빈 문자열("")로 채워 길이를 맞춥니다.
     *
     * @param row 원본 행 데이터
     * @param colCount 목표 컬럼 수
     * @return 컬럼 수에 맞게 패딩된 행 배열
     */
    private String[] padRow(String[] row, int colCount) {
        if (row.length == colCount) return row;
        String[] padded = new String[colCount];
        System.arraycopy(row, 0, padded, 0, row.length);
        for (int i = row.length; i < colCount; i++) padded[i] = "";
        return padded;
    }
}
