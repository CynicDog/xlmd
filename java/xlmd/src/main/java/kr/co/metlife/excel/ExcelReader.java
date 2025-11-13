package kr.co.metlife.excel;

import kr.co.metlife.excel.model.SheetData;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

/**
 * 사용자가 입력한 Excel 테이블 데이터를 읽어 {@link SheetData} 모델로 변환하는 클래스입니다.
 */
public class ExcelReader {

    /**
     * 사용자 입력(stdin)으로부터 여러 줄의 탭 또는 콤마 구분 데이터를 읽습니다.
     * 빈 줄을 만나면 입력을 종료합니다.
     *
     * @return 입력된 데이터를 포함하는 {@link SheetData} 리스트
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
