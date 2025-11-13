package kr.co.metlife.excel;

import kr.co.metlife.excel.model.SheetData;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.io.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Minimal XLSX writer using only core Java (ZIP + StAX XML).
 * Produces the same structure as the working Go version.
 */
public class ExcelWriter {

    private static final String NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";

    private final String filePath;
    private final List<SheetData> sheets;

    private List<String> sharedStrings;
    private Map<String, Integer> stringIndexMap;
    private XMLOutputFactory xmlFactory;
    private ZipOutputStream zipOut;

    /**
     * Creates a new {@code ExcelWriter} for writing sheet data to an XLSX file.
     *
     * @param filePath The output file path where the XLSX will be written.
     * @param sheets The list of {@link SheetData} objects to include in the workbook.
     */
    public ExcelWriter(String filePath, List<SheetData> sheets) {
        this.filePath = filePath;
        this.sheets = sheets;
        this.xmlFactory = XMLOutputFactory.newInstance();
    }

    /**
     * Writes the XLSX file to disk, generating all required XML parts
     * and packaging them into a ZIP archive.
     *
     * @throws IOException If an error occurs while creating or writing the XLSX file.
     */
    public void write() throws IOException {
        buildSharedStringTable();

        try (FileOutputStream fos = new FileOutputStream(filePath);
             BufferedOutputStream bos = new BufferedOutputStream(fos);
             ZipOutputStream zos = new ZipOutputStream(bos)) {

            this.zipOut = zos;

            // [Content_Types].xml
            writeContentTypes();

            // _rels/.rels
            writeRootRelationships();

            // xl/workbook.xml
            writeWorkbookXML();

            // xl/_rels/workbook.xml.rels
            writeWorkbookRelationships();

            // xl/styles.xml
            writeStylesXML();

            // xl/sharedStrings.xml
            writeSharedStringsXML();

            // Each worksheet
            for (int i = 0; i < sheets.size(); i++) {
                writeWorksheetXML(sheets.get(i), i + 1);
            }

            zipOut.finish();

        } catch (Exception e) {
            throw new IOException("Failed to write Excel file: " + e.getMessage(), e);
        }
    }

    /**
     * Builds the shared string table used by the workbook.
     * Each unique string is assigned an index for reference by cells.
     */
    private void buildSharedStringTable() {
        this.sharedStrings = new ArrayList<>();
        this.stringIndexMap = new HashMap<>();

        for (SheetData sheet : sheets) {
            List<String[]> rows = sheet.getRows();
            if (rows == null) continue;

            for (String[] row : rows) {
                if (row == null) continue;

                // Loop through each cell value in the String[]
                for (String value : row) {
                    if (value != null && !value.isEmpty() && !stringIndexMap.containsKey(value)) {
                        stringIndexMap.put(value, sharedStrings.size());
                        sharedStrings.add(value);
                    }
                }
            }
        }
    }

    /**
     * Creates an {@link XMLStreamWriter} for a new ZIP entry within the XLSX archive.
     *
     * @param entry The path of the ZIP entry to create (e.g., "xl/workbook.xml").
     * @return A new {@link XMLStreamWriter} for writing XML content to the entry.
     * @throws IOException If an I/O error occurs while creating the ZIP entry.
     * @throws XMLStreamException If an XML streaming error occurs during writer creation.
     */
    private XMLStreamWriter createXMLWriter(String entry) throws IOException, XMLStreamException {
        zipOut.putNextEntry(new ZipEntry(entry));
        XMLStreamWriter writer = xmlFactory.createXMLStreamWriter(zipOut, "UTF-8");
        writer.writeStartDocument("UTF-8", "1.0");
        return writer;
    }

    /**
     * Closes the given XML stream writer and finalizes the current ZIP entry.
     *
     * @param writer The {@link XMLStreamWriter} to close.
     * @throws XMLStreamException If an XML streaming error occurs during closing.
     * @throws IOException If an I/O error occurs while closing the ZIP entry.
     */
    private void closeXMLWriter(XMLStreamWriter writer) throws XMLStreamException, IOException {
        writer.writeEndDocument();
        writer.flush();
        writer.close();
        zipOut.closeEntry();
    }

    /**
     * Writes the content types XML file ([Content_Types].xml),
     * defining MIME types for all parts of the Excel package.
     *
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeContentTypes() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("[Content_Types].xml");
        w.writeStartElement("Types");
        w.writeDefaultNamespace(NS_CT);

        w.writeStartElement("Default");
        w.writeAttribute("Extension", "rels");
        w.writeAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        w.writeEndElement();

        w.writeStartElement("Default");
        w.writeAttribute("Extension", "xml");
        w.writeAttribute("ContentType", "application/xml");
        w.writeEndElement();

        w.writeStartElement("Override");
        w.writeAttribute("PartName", "/xl/workbook.xml");
        w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        w.writeEndElement();

        w.writeStartElement("Override");
        w.writeAttribute("PartName", "/xl/styles.xml");
        w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        w.writeEndElement();

        w.writeStartElement("Override");
        w.writeAttribute("PartName", "/xl/sharedStrings.xml");
        w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        w.writeEndElement();

        for (int i = 0; i < sheets.size(); i++) {
            w.writeStartElement("Override");
            w.writeAttribute("PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml");
            w.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            w.writeEndElement();
        }

        w.writeEndElement();
        closeXMLWriter(w);
    }

    /**
     * Writes the root relationships XML file (_rels/.rels),
     * linking the package to the main workbook document.
     *
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeRootRelationships() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("_rels/.rels");
        w.writeStartElement("Relationships");
        w.writeDefaultNamespace(NS_REL);

        w.writeStartElement("Relationship");
        w.writeAttribute("Id", "rId1");
        w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        w.writeAttribute("Target", "xl/workbook.xml");
        w.writeEndElement();

        w.writeEndElement();
        closeXMLWriter(w);
    }

    /**
     * Writes the workbook relationships XML file (xl/_rels/workbook.xml.rels),
     * defining links between the workbook and its sheets, styles, and shared strings.
     *
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeWorkbookRelationships() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("xl/_rels/workbook.xml.rels");
        w.writeStartElement("Relationships");
        w.writeDefaultNamespace(NS_REL);

        for (int i = 0; i < sheets.size(); i++) {
            w.writeStartElement("Relationship");
            w.writeAttribute("Id", "rId" + (i + 1));
            w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            w.writeAttribute("Target", "worksheets/sheet" + (i + 1) + ".xml");
            w.writeEndElement();
        }

        int stylesId = sheets.size() + 1;
        w.writeStartElement("Relationship");
        w.writeAttribute("Id", "rId" + stylesId);
        w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        w.writeAttribute("Target", "styles.xml");
        w.writeEndElement();

        int sharedId = sheets.size() + 2;
        w.writeStartElement("Relationship");
        w.writeAttribute("Id", "rId" + sharedId);
        w.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
        w.writeAttribute("Target", "sharedStrings.xml");
        w.writeEndElement();

        w.writeEndElement();
        closeXMLWriter(w);
    }

    /**
     * Writes the workbook XML file (xl/workbook.xml), defining all sheets and their relationships.
     *
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeWorkbookXML() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("xl/workbook.xml");
        w.writeStartElement("workbook");
        w.writeDefaultNamespace(NS_MAIN);

        w.writeStartElement("sheets");
        for (int i = 0; i < sheets.size(); i++) {
            w.writeStartElement("sheet");
            w.writeAttribute("name", sheets.get(i).getName());
            w.writeAttribute("sheetId", String.valueOf(i + 1));
            w.writeAttribute("xmlns:relationships", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            w.writeAttribute("relationships:id", "rId" + (i + 1));
            w.writeEndElement();
        }
        w.writeEndElement(); // </sheets>

        w.writeEndElement(); // </workbook>
        closeXMLWriter(w);
    }

    /**
     * Writes a minimal styles XML file (xl/styles.xml) required by Excel for workbook formatting.
     *
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeStylesXML() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("xl/styles.xml");
        w.writeStartElement("styleSheet");
        w.writeDefaultNamespace(NS_MAIN);
        w.writeEndElement();
        closeXMLWriter(w);
    }

    /**
     * Writes the shared strings XML file (xl/sharedStrings.xml), which stores all unique text values.
     *
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeSharedStringsXML() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("xl/sharedStrings.xml");
        w.writeStartElement("sst");
        w.writeDefaultNamespace(NS_MAIN);
        w.writeAttribute("count", String.valueOf(sharedStrings.size()));
        w.writeAttribute("uniqueCount", String.valueOf(sharedStrings.size()));

        for (String s : sharedStrings) {
            w.writeStartElement("si");
            w.writeStartElement("t");
            w.writeCharacters(s);
            w.writeEndElement();
            w.writeEndElement();
        }

        w.writeEndElement();
        closeXMLWriter(w);
    }

    /**
     * Writes a worksheet XML file (xl/worksheets/sheetX.xml) for the given sheet.
     * Each cell is written using shared string references.
     *
     * @param sheet The {@link SheetData} object containing the worksheet data.
     * @param sheetId The 1-based index of the sheet, used to name the XML file.
     * @throws IOException If an I/O error occurs while writing to the ZIP output.
     * @throws XMLStreamException If an XML streaming error occurs during writing.
     */
    private void writeWorksheetXML(SheetData sheet, int sheetId) throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("xl/worksheets/sheet" + sheetId + ".xml");
        w.writeStartElement("worksheet");
        w.writeDefaultNamespace(NS_MAIN);

        w.writeStartElement("sheetData");

        List<String[]> rows = sheet.getRows();

        if (rows != null) {
            for (int rIdx = 0; rIdx < rows.size(); rIdx++) {
                String[] row = rows.get(rIdx);

                if (row == null || row.length == 0) continue;
                int rowNum = rIdx + 1;

                w.writeStartElement("row");
                w.writeAttribute("r", String.valueOf(rowNum));

                // Iterate over the String array
                for (int cIdx = 0; cIdx < row.length; cIdx++) {
                    String value = row[cIdx];

                    if (value == null || value.isEmpty()) continue;
                    String colName = indexToColRef(cIdx);
                    int index = stringIndexMap.getOrDefault(value, -1);
                    if (index < 0) continue;

                    w.writeStartElement("c");
                    w.writeAttribute("r", colName + rowNum);
                    w.writeAttribute("t", "s");

                    w.writeStartElement("v");
                    w.writeCharacters(String.valueOf(index));
                    w.writeEndElement(); // </v>

                    w.writeEndElement(); // </c>
                }

                w.writeEndElement(); // </row>
            }
        }
        w.writeEndElement(); // </sheetData>
        w.writeEndElement(); // </worksheet>
        closeXMLWriter(w);
    }

    /**
     * Converts a zero-based column index to an Excel-style column letter.
     * For example, 0 → "A", 25 → "Z", 26 → "AA".
     *
     * @param col The zero-based column index.
     * @return The corresponding Excel-style column letter.
     */
    private String indexToColRef(int col) {
        StringBuilder sb = new StringBuilder();
        while (col >= 0) {
            sb.insert(0, (char) ('A' + (col % 26)));
            col = (col / 26) - 1;
        }
        return sb.toString();
    }
}
