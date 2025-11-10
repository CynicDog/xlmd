package kr.co.metlife.excel.model;

/**
 * Helper class to temporarily hold raw cell data read from xl/worksheets/sheetX.xml
 * before it is processed and mapped into the simple List<List<String>> structure in SheetData.
 *
 * Corresponds to the <c> element in the XML: <c r="A1" t="s">...</c>
 */
public class RawCell {
    // Cell reference, e.g., "A1", "C5" (from 'r' attribute)
    private String ref;

    // Cell data type, e.g., "s" for shared string, "n" for number (from 't' attribute)
    private String type;

    // The cell's raw value, or the index if type is "s" (content of <v> tag)
    private String value;

    public RawCell(String ref, String type, String value) {
        this.ref = ref;
        this.type = type;
        this.value = value;
    }

    // Getters for XML reading logic
    public String getRef() {
        return ref;
    }

    public String getType() {
        return type;
    }

    public String getValue() {
        return value;
    }


}