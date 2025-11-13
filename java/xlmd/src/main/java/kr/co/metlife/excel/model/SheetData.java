package kr.co.metlife.excel.model;
/*
 * Excel Package Types
 *
 * This file defines the XML structures and core data models used by the
 * `xlmd/excel` package for reading and writing .xlsx files. The .xlsx
 * format is actually a ZIP archive containing a hierarchy of XML parts,
 * as illustrated below:
 *
 * example.xlsx                  : actually a ZIP archive
 * │
 * ├── [Content_Types].xml       : MIME/content type definitions for all parts
 * │
 * ├── _rels/                    : package-level relationships
 * │   └── .rels                 : links root → /xl/workbook.xml
 * │
 * ├── docProps/                 : document metadata
 * │   ├── app.xml               : application info (Excel version, total sheets, etc.)
 * │   └── core.xml              : author, title, created/modified dates
 * │
 * └── xl/                       : main workbook content
 *     │
 *     ├── workbook.xml          : workbook definition (sheet list, names, order)
 *     ├── _rels/                : workbook-level relationships
 *     │   └── workbook.xml.rels : maps sheets, sharedStrings, styles by rId
 *     │
 *     ├── sharedStrings.xml     : table of all unique text strings (string pool)
 *     ├── styles.xml            : cell styles; fonts, fills, number formats
 *     ├── theme/                : color & theme info
 *     │   └── theme1.xml
 *     │
 *     └── worksheets/           : actual sheet data (one XML per sheet)
 *         ├── sheet1.xml        : e.g., “Sheet1”
 *         ├── sheet2.xml        : e.g., “Sheet2”
 *         └── sheetN.xml        : and so on...
 *
 * These structures are designed to provide a clean separation between
 * raw XML serialization (used by the XLSX format) and the higher-level
 * `SheetData` abstraction used throughout xlmd.
 *
 */

import java.util.List;

/**
 * 하나의 시트를 표현하는 데이터 모델입니다.
 * 시트 이름과 여러 행(row)을 포함하며, 각 행은 문자열 배열로 표현됩니다.
 */
public class SheetData {
    private String name;
    private List<String[]> rows;

    /**
     * 시트 데이터를 생성합니다.
     *
     * @param name 시트 이름
     * @param rows 시트의 행 데이터 리스트, 각 행은 문자열 배열
     */
    public SheetData(String name, List<String[]> rows) {
        this.name = name;
        this.rows = rows;
    }

    public String getName() {
        return name;
    }

    public List<String[]> getRows() {
        return rows;
    }

    /**
     * 시트 내에서 가장 많은 열(column) 수를 반환합니다.
     *
     * @return 최대 열 수
     */
    public int getMaxColumnCount() {
        return rows.stream().mapToInt(r -> r.length).max().orElse(0);
    }
}
