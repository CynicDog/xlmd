package kr.co.metlife.markdown;

import kr.co.metlife.excel.model.SheetData;

import java.util.*;
import java.util.regex.*;

/**
 * CLI 입력으로부터 Markdown 테이블 문자열을 읽어 SheetData로 변환합니다.
 */
public class MarkdownReader {

    private static final Pattern SEPARATOR_ROW_PATTERN = Pattern.compile("^\\|(\\s*[-:]+\\s*\\|)+\\s*$");

    /**
     * 주어진 Markdown 문자열을 읽어 하나의 시트(Sheet1)로 변환합니다.
     *
     * - 내부적으로 parseTable을 사용하여 테이블 행을 추출합니다.
     * - 각 행은 문자열 배열로 변환되어 SheetData에 저장됩니다.
     *
     * @param markdown Markdown 테이블 문자열
     * @return 하나의 SheetData를 담은 리스트
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

    /**
     * 주어진 Markdown 텍스트에서 테이블을 파싱하여
     * 각 행을 리스트 형태로 반환합니다.
     *
     * - 각 행은 | 구분자로 나누어져 있으며, 앞뒤 공백은 제거됩니다.
     * - 구분 행(---)은 무시됩니다.
     * - 부족한 컬럼은 빈 문자열로 패딩되며, 초과 컬럼은 잘립니다.
     *
     * @param text Markdown 테이블 텍스트
     * @return 테이블의 각 행을 담은 리스트
     */
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
