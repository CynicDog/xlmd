package kr.co.metlife.excel;

import kr.co.metlife.excel.model.RawCell;
import kr.co.metlife.excel.model.SheetData;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

/**
 * Reads an .xlsx file using only core Java libraries (java.util.zip and StAX).
 * This class performs the procedural reading of the XML parts.
 */
public class ExcelReader {

    private final String filePath;

    public ExcelReader(String filePath) {
        this.filePath = filePath;
    }

    /**
     * Helper class to ensure the ZipInputStream and its underlying FileInputStream
     * are properly closed after a single logical read operation (used by getEntryInputStream).
     */
    private static class StreamCloser extends InputStream {
        private final ZipInputStream zis;
        private final FileInputStream fis;

        public StreamCloser(ZipInputStream zis, FileInputStream fis) {
            this.zis = zis;
            this.fis = fis;
        }

        @Override
        public int read() throws IOException {
            return zis.read();
        }

        @Override
        public int read(byte[] b, int off, int len) throws IOException {
            return zis.read(b, off, len);
        }

        @Override
        public void close() throws IOException {
            // Close both the ZIS and the FIS used for this specific entry read
            zis.close();
            fis.close();
        }
    }

    /**
     * Retrieves an input stream for a specific file entry within the XLSX ZIP package.
     * A new {@link FileInputStream} is opened for each call to allow independent access.
     *
     * @param path The internal ZIP path of the entry (e.g., "xl/workbook.xml").
     * @return An {@link InputStream} for the specified entry, or {@code null} if not found.
     * @throws IOException If an I/O error occurs while accessing the ZIP file.
     */
    private InputStream getEntryInputStream(String path) throws IOException {
        // We need a fresh FileInputStream for each part, as ZipInputStream is single-pass.
        FileInputStream fis = new FileInputStream(filePath);
        ZipInputStream zis = new ZipInputStream(fis);
        ZipEntry entry;

        while ((entry = zis.getNextEntry()) != null) {
            if (entry.getName().equals(path)) {
                // Return a wrapped stream that handles closing both the ZIS and FIS
                return new StreamCloser(zis, fis);
            }
        }
        // If entry not found, close the streams and return null
        zis.close();
        fis.close();
        return null;
    }

    /**
     * Reads the XLSX file and reconstructs its sheets into a list of {@link SheetData} objects.
     * This includes parsing shared strings, workbook structure, relationships, and sheet data.
     *
     * @return A list of {@link SheetData} objects representing all sheets in the workbook.
     * @throws IOException If an I/O error occurs while reading from the XLSX file.
     * @throws XMLStreamException If an XML parsing error occurs during reading.
     */
    public List<SheetData> read() throws IOException, XMLStreamException {
        List<SheetData> sheets = new ArrayList<>();
        List<String> sharedStrings = new ArrayList<>();
        Map<String, String> sheetRidsToNames = new HashMap<>();
        Map<String, String> relIdToTargetPath = new HashMap<>();

        // Read Shared Strings (xl/sharedStrings.xml)
        try (InputStream ssStream = getEntryInputStream("xl/sharedStrings.xml")) {
            if (ssStream != null) {
                sharedStrings = readSharedStringsXML(ssStream);
            }
        }

        // Read Workbook structure (xl/workbook.xml) and Relationships (xl/_rels/workbook.xml.rels)
        try (InputStream wbStream = getEntryInputStream("xl/workbook.xml")) {
            if (wbStream != null) {
                sheetRidsToNames = readWorkbookXML(wbStream);
            }
        }

        try (InputStream relStream = getEntryInputStream("xl/_rels/workbook.xml.rels")) {
            if (relStream != null) {
                relIdToTargetPath = readWorkbookRels(relStream);
            }
        }

        // Read Sheet Data (xl/worksheets/sheetX.xml)
        // Iterate through sheets discovered in workbook.xml
        for (Map.Entry<String, String> entryMap : sheetRidsToNames.entrySet()) {
            String rId = entryMap.getKey();
            String sheetName = entryMap.getValue();

            String targetPath = relIdToTargetPath.get(rId);

            if (targetPath == null) {
                System.err.printf("Warning: Could not find target relationship for sheet '%s' with rId '%s'. Skipping.%n", sheetName, rId);
                continue;
            }

            // The full path inside the ZIP is "xl/" + targetPath
            String fullPath = "xl/" + targetPath;

            // FIX: Use a try-catch-continue block for individual sheet reading to skip corrupted sheets.
            try (InputStream sheetStream = getEntryInputStream(fullPath)) {
                if (sheetStream == null) {
                    // This handles the 'ZIP entry not found' case (e.g., sheet4.xml is missing)
                    System.err.printf("Warning: Could not read sheet file %s: ZIP entry not found: %s%n", fullPath, fullPath);
                    continue; // Skip this sheet and proceed to the next one
                }

                // Read the actual sheet data
                SheetData sheetData = readSheetXML(sheetStream, sheetName, sharedStrings);
                sheets.add(sheetData);

            } catch (IOException | XMLStreamException e) {
                // Catch any IO or XML parsing errors specific to this sheet and continue
                System.err.printf("Error reading sheet '%s' at %s: %s%n", sheetName, fullPath, e.getMessage());
                continue; // Go to the next sheet
            }
        }

        return sheets;
    }

    /**
     * Reads the shared strings XML file and extracts all string values.
     * Each entry corresponds to a unique shared string used in the workbook.
     *
     * @param ssStream The input stream of the shared strings XML file.
     * @return A list of all shared string values in order of appearance.
     * @throws XMLStreamException If an XML parsing error occurs during reading.
     */
    private List<String> readSharedStringsXML(InputStream ssStream) throws XMLStreamException {
        List<String> strings = new ArrayList<>();
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLEventReader reader = factory.createXMLEventReader(ssStream);
        String currentText = "";

        while (reader.hasNext()) {
            XMLEvent event = reader.nextEvent();
            if (event.isStartElement()) {
                StartElement startElement = event.asStartElement();
                String name = startElement.getName().getLocalPart();

                if (name.equals("t")) { // <t> tag contains the actual string
                    event = reader.nextEvent();
                    if (event.isCharacters()) {
                        currentText = event.asCharacters().getData();
                    }
                }
            } else if (event.isEndElement()) {
                EndElement endElement = event.asEndElement();
                if (endElement.getName().getLocalPart().equals("si")) { // </si> end
                    strings.add(currentText);
                    currentText = ""; // Reset for the next string
                }
            }
        }
        reader.close();
        return strings;
    }

    /**
     * Reads the workbook XML file and maps relationship IDs to sheet names.
     *
     * @param wbStream The input stream of the workbook XML file.
     * @return A map of relationship IDs to their corresponding sheet names.
     * @throws XMLStreamException If an XML parsing error occurs during reading.
     */
    private Map<String, String> readWorkbookXML(InputStream wbStream) throws XMLStreamException {
        Map<String, String> sheetRidsToNames = new HashMap<>();
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLEventReader reader = factory.createXMLEventReader(wbStream);

        while (reader.hasNext()) {
            XMLEvent event = reader.nextEvent();
            if (event.isStartElement()) {
                StartElement startElement = event.asStartElement();
                if (startElement.getName().getLocalPart().equals("sheet")) {
                    String name = getAttributeValue(startElement, "name", null);
                    // Relationship ID uses the specific relationship namespace
                    String rId = getAttributeValue(startElement, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    if (name != null && rId != null) {
                        sheetRidsToNames.put(rId, name);
                    }
                }
            }
        }
        reader.close();
        return sheetRidsToNames;
    }

    /**
     * Reads the workbook relationships XML file and maps relationship IDs to worksheet file paths.
     *
     * @param relStream The input stream of the workbook relationships XML file.
     * @return A map of relationship IDs to their corresponding worksheet target paths.
     * @throws XMLStreamException If an XML parsing error occurs during reading.
     */
    private Map<String, String> readWorkbookRels(InputStream relStream) throws XMLStreamException {
        Map<String, String> relIdToTargetPath = new HashMap<>();
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLEventReader reader = factory.createXMLEventReader(relStream);

        while (reader.hasNext()) {
            XMLEvent event = reader.nextEvent();
            if (event.isStartElement()) {
                StartElement startElement = event.asStartElement();
                if (startElement.getName().getLocalPart().equals("Relationship")) {
                    String rId = getAttributeValue(startElement, "Id", null);
                    String target = getAttributeValue(startElement, "Target", null);
                    // Only store relationships to worksheet files (they are relative paths like "worksheets/sheet1.xml")
                    if (rId != null && target != null && target.startsWith("worksheets/")) {
                        relIdToTargetPath.put(rId, target);
                    }
                }
            }
        }
        reader.close();
        return relIdToTargetPath;
    }

    /**
     * Reads a worksheet XML file and converts it into a {@link SheetData} object.
     * Resolves shared string references and reconstructs each row and cell value.
     *
     * @param sheetStream The input stream of the worksheet XML file.
     * @param sheetName The name of the sheet being read.
     * @param sharedStrings The list of shared string values used for text cell resolution.
     * @return A {@link SheetData} object containing the parsed sheet data.
     * @throws XMLStreamException If an XML parsing error occurs during reading.
     */
    private SheetData readSheetXML(InputStream sheetStream, String sheetName, List<String> sharedStrings) throws XMLStreamException {
        SheetData sheet = new SheetData(sheetName);
        XMLInputFactory factory = XMLInputFactory.newInstance();
        XMLEventReader reader = factory.createXMLEventReader(sheetStream);

        List<RawCell> currentRowCells = new ArrayList<>();

        while (reader.hasNext()) {
            XMLEvent event = reader.nextEvent();
            if (event.isStartElement()) {
                StartElement startElement = event.asStartElement();
                String name = startElement.getName().getLocalPart();

                if (name.equals("row")) {
                    // Start of a new row, clear previous cells
                    currentRowCells.clear();

                } else if (name.equals("c")) { // Cell
                    String ref = getAttributeValue(startElement, "r", null);
                    String type = getAttributeValue(startElement, "t", null);

                    // Look for the <v> value tag immediately following the <c> tag
                    String val = "";
                    // Advance past subsequent events until <v> start tag or </c> end tag
                    while (reader.hasNext()) {
                        XMLEvent innerEvent = reader.nextEvent();
                        if (innerEvent.isStartElement() && innerEvent.asStartElement().getName().getLocalPart().equals("v")) {
                            // Read value within <v>
                            innerEvent = reader.nextEvent();
                            if (innerEvent.isCharacters()) {
                                val = innerEvent.asCharacters().getData();
                            }
                            // Advance past </v>
                            while (reader.hasNext() && !reader.nextEvent().isEndElement()) {}

                        } else if (innerEvent.isEndElement() && innerEvent.asEndElement().getName().getLocalPart().equals("c")) {
                            break; // End of cell
                        }
                    }

                    if (ref != null && val != null) {
                        currentRowCells.add(new RawCell(ref, type, val));
                    }

                }
            } else if (event.isEndElement()) {
                EndElement endElement = event.asEndElement();
                if (endElement.getName().getLocalPart().equals("row") && !currentRowCells.isEmpty()) {
                    // End of row: Process and add to sheet data

                    // Find the max column index required for this row based on cell references
                    int maxColIndex = 0;
                    for (RawCell cell : currentRowCells) {
                        int colIndex = colRefToIndex(cell.getRef());
                        if (colIndex > maxColIndex) {
                            maxColIndex = colIndex;
                        }
                    }

                    List<String> rowVals = new ArrayList<>(maxColIndex + 1);
                    // Initialize the row with empty strings
                    for (int i = 0; i <= maxColIndex; i++) {
                        rowVals.add("");
                    }

                    // Populate the row slice by placing values at the correct column index.
                    for (RawCell cell : currentRowCells) {
                        int colIndex = colRefToIndex(cell.getRef());
                        String v = cell.getValue();

                        // Resolve shared strings if type="s"
                        if ("s".equals(cell.getType())) {
                            try {
                                int idx = Integer.parseInt(v);
                                if (idx >= 0 && idx < sharedStrings.size()) {
                                    v = sharedStrings.get(idx);
                                } else {
                                    v = "";
                                }
                            } catch (NumberFormatException e) {
                                v = "";
                            }
                        }

                        rowVals.set(colIndex, v);
                    }

                    sheet.addRow(rowVals);
                }
            }
        }
        reader.close();
        return sheet;
    }


    /**
     * Retrieves the value of an attribute from an XML start element.
     *
     * @param element The {@link StartElement} to read from.
     * @param localName The local name of the attribute.
     * @param namespaceUri The namespace URI of the attribute, or {@code null} if none.
     * @return The attribute value, or {@code null} if the attribute is not present.
     */
    private String getAttributeValue(StartElement element, String localName, String namespaceUri) {
        Attribute attribute = null;
        if (namespaceUri == null) {
            attribute = element.getAttributeByName(
                    new javax.xml.namespace.QName(localName));
        } else {
            attribute = element.getAttributeByName(
                    new javax.xml.namespace.QName(namespaceUri, localName));
        }

        return (attribute != null) ? attribute.getValue() : null;
    }

    /**
     * Converts an Excel-style column reference (e.g., "A", "AB") to a zero-based column index.
     *
     * @param ref The Excel column reference string.
     * @return The zero-based column index corresponding to the given reference.
     */
    private int colRefToIndex(String ref) {
        String colStr = "";
        for (char c : ref.toCharArray()) {
            if (c >= 'A' && c <= 'Z') {
                colStr += c;
            } else {
                break;
            }
        }

        int index = 0;
        for (int i = 0; i < colStr.length(); i++) {
            char charVal = colStr.charAt(i);
            int power = colStr.length() - i - 1;
            // Char value 'A' is 65. ('A' - 'A' + 1) = 1.
            // Power of 26: 26^0 for the last letter, 26^1 for the second to last, etc.
            index += (int) ( (charVal - 'A' + 1) * Math.pow(26, power) );
        }
        // Convert 1-based index (1=A) to 0-based index (0=A)
        return index - 1;
    }
}