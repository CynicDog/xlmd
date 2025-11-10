package kr.co.metlife.excel.model;

/*
 *
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

import java.util.ArrayList;
import java.util.List;

public class SheetData {

    private String name;

    // List of rows, where each row is a List of cell values (strings).
    private List<List<String>> rows;

    /**
     * For building a sheet incrementally (your original version).
     */
    public SheetData(String name) {
        this.name = name;
        this.rows = new ArrayList<>();
    }

    /**
     * For building a sheet with all data at once (used by MarkdownReader).
     */
    public SheetData(String name, List<List<String>> rows) {
        this.name = name;
        this.rows = rows;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<List<String>> getRows() {
        return rows;
    }

    public void addRow(List<String> row) {
        this.rows.add(row);
    }
}