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
 * 순수 Java(ZIP + StAX XML)만 사용하여 OpenXML 규격을 준수하는 XLSX 파일을 작성하는 클래스입니다.
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
     * XLSX 파일로 시트 데이터를 작성하기 위한 새로운 {@code ExcelWriter} 객체를 생성합니다.
     *
     * @param filePath XLSX 파일이 작성될 출력 경로
     * @param sheets 워크북에 포함될 {@link SheetData} 객체들의 목록
     */
    public ExcelWriter(String filePath, List<SheetData> sheets) {
        this.filePath = filePath;
        this.sheets = sheets;
        this.xmlFactory = XMLOutputFactory.newInstance();
    }

    /**
     * XLSX 파일을 디스크에 작성합니다.
     * 필요한 모든 XML 파트를 생성하고, 이를 ZIP 파일로 패키징합니다.
     *
     * @throws IOException XLSX 파일 생성 또는 작성 중 오류가 발생한 경우
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
     * 워크북에서 사용할 공유 문자열 테이블을 구성합니다.
     * 각 고유 문자열은 셀에서 참조할 수 있도록 인덱스를 할당받습니다.
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
     * XLSX 아카이브 내 새 ZIP 엔트리에 대한 {@link XMLStreamWriter}를 생성합니다.
     *
     * @param entry 생성할 ZIP 엔트리 경로 (예: "xl/workbook.xml")
     * @return 해당 엔트리에 XML 내용을 작성할 새로운 {@link XMLStreamWriter}
     * @throws IOException ZIP 엔트리를 생성하는 동안 I/O 오류가 발생한 경우
     * @throws XMLStreamException 작성기 생성 중 XML 스트리밍 오류가 발생한 경우
     */
    private XMLStreamWriter createXMLWriter(String entry) throws IOException, XMLStreamException {
        zipOut.putNextEntry(new ZipEntry(entry));
        XMLStreamWriter writer = xmlFactory.createXMLStreamWriter(zipOut, "UTF-8");
        writer.writeStartDocument("UTF-8", "1.0");
        return writer;
    }

    /**
     * 지정된 XML 스트림 작성기를 닫고 현재 ZIP 엔트리를 완료합니다.
     *
     * @param writer 닫을 {@link XMLStreamWriter}
     * @throws XMLStreamException 닫는 동안 XML 스트리밍 오류가 발생한 경우
     * @throws IOException ZIP 엔트리를 닫는 동안 I/O 오류가 발생한 경우
     */
    private void closeXMLWriter(XMLStreamWriter writer) throws XMLStreamException, IOException {
        writer.writeEndDocument();
        writer.flush();
        writer.close();
        zipOut.closeEntry();
    }

    /**
     * 콘텐츠 유형 XML 파일([Content_Types].xml)을 작성합니다.
     * Excel 패키지의 각 파트에 대한 MIME 타입을 정의합니다.
     *
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
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
     * 루트 관계 XML 파일(_rels/.rels)을 작성합니다.
     * 이 파일은 패키지와 메인 워크북 문서(xl/workbook.xml) 간의 연결을 정의합니다.
     *
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
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
     * 워크북과 각 시트, 스타일, 공유 문자열 간의 관계를 정의하는
     * 워크북 관계 XML 파일(xl/_rels/workbook.xml.rels)을 작성합니다.
     *
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
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
     * 모든 시트와 그 관계를 정의하는 워크북 XML 파일(xl/workbook.xml)을 작성합니다.
     *
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
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
     * Excel 워크북 형식 작성을 위해 필요한 최소한의 스타일 XML 파일(xl/styles.xml)을 작성합니다.
     *
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
     */
    private void writeStylesXML() throws IOException, XMLStreamException {
        XMLStreamWriter w = createXMLWriter("xl/styles.xml");
        w.writeStartElement("styleSheet");
        w.writeDefaultNamespace(NS_MAIN);
        w.writeEndElement();
        closeXMLWriter(w);
    }

    /**
     * 모든 고유 문자열 값을 저장하는 shared strings XML 파일(xl/sharedStrings.xml)을 작성합니다.
     *
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
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
     * 주어진 시트 데이터를 기반으로 워크시트 XML 파일(xl/worksheets/sheetX.xml)을 작성합니다.
     * 각 셀은 shared string 참조를 사용하여 기록됩니다.
     *
     * @param sheet 워크시트 데이터를 포함하는 {@link SheetData} 객체
     * @param sheetId 시트의 1 기반 인덱스, XML 파일 이름에 사용
     * @throws IOException ZIP 출력 중 I/O 오류가 발생할 경우
     * @throws XMLStreamException XML 스트리밍 작성 중 오류가 발생할 경우
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
     * 0 기반 열 인덱스를 Excel 스타일 열 문자로 변환합니다.
     * 예: 0 → "A", 25 → "Z", 26 → "AA"
     *
     * @param col 0 기반 열 인덱스
     * @return 해당하는 Excel 열 문자
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
