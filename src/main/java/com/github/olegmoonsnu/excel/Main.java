package com.github.olegmoonsnu.excel;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URI;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.StringJoiner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipFile;
import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

public class Main {

    // Настройки. Номера листов и строк начинаются с 1, как в Excel.
    private static final String EXCEL_FILE = "input.xlsx";
    private static final int SHEET_NUMBER = 1;
    private static final int FIRST_DATA_ROW = 2;
    private static final String COLUMN = "B";
    private static final int COLUMN_OFFSET = 0; // -1: A, 0: B, +1: C
    private static final boolean SPLIT_BY_COMMA = true;
    private static final boolean SPLIT_BY_NEWLINE = true;
    private static final ArrayFormat OUTPUT_FORMAT = ArrayFormat.JAVA;

    public static void main(String[] args) {
        if (args.length > 1) {
            System.err.println("Передайте один аргумент с путём к Excel-файлу.");
            System.exit(1);
        }

        try {
            Path file = Path.of(args.length == 1 ? args[0] : EXCEL_FILE);
            if (file.toString().toLowerCase(Locale.ROOT).endsWith(".xls")) {
                throw new IllegalArgumentException("Сохраните файл .xls в Excel в формате .xlsx.");
            }
            List<String> values = XlsxColumnReader.read(file, SHEET_NUMBER, FIRST_DATA_ROW,
                    COLUMN, COLUMN_OFFSET, createSeparator());
            System.out.println(OUTPUT_FORMAT.format(values));
        } catch (IOException | IllegalArgumentException exception) {
            System.err.println("Ошибка чтения Excel. " + exception.getMessage());
            System.exit(1);
        }
    }

    private static Pattern createSeparator() {
        List<String> separators = new ArrayList<>();
        if (SPLIT_BY_COMMA) {
            separators.add(",");
        }
        if (SPLIT_BY_NEWLINE) {
            separators.add("\\R");
        }
        return separators.isEmpty() ? null : Pattern.compile(String.join("|", separators));
    }
}

enum ArrayFormat {
    JAVA("new String[] {", "}", '"'),
    GROOVY("[", "] as String[]", '\'');

    private final String prefix;
    private final String suffix;
    private final char quote;

    ArrayFormat(String prefix, String suffix, char quote) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.quote = quote;
    }

    String format(List<String> values) {
        StringJoiner array = new StringJoiner(", ", prefix, suffix);
        for (String value : values) {
            array.add(quote(value));
        }
        return array.toString();
    }

    private String quote(String value) {
        StringBuilder result = new StringBuilder().append(quote);
        for (int i = 0; i < value.length(); i++) {
            char character = value.charAt(i);
            switch (character) {
                case '\\' -> result.append("\\\\");
                case '\n' -> result.append("\\n");
                case '\r' -> result.append("\\r");
                case '\t' -> result.append("\\t");
                case '\b' -> result.append("\\b");
                case '\f' -> result.append("\\f");
                default -> {
                    if (character == quote) {
                        result.append('\\').append(character);
                    } else if (Character.isISOControl(character)
                            || character == '\u2028' || character == '\u2029') {
                        result.append(String.format(Locale.ROOT, "\\u%04x", (int) character));
                    } else {
                        result.append(character);
                    }
                }
            }
        }
        return result.append(quote).toString();
    }
}

final class XlsxColumnReader {

    private static final int MAX_COLUMNS = 16_384;
    private static final int MAX_ROWS = 1_048_576;
    private static final String RELATIONSHIP_NS =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String STRICT_RELATIONSHIP_NS =
            "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private static final Pattern EXCEL_ESCAPE = Pattern.compile("_x([0-9a-fA-F]{4})_");

    private XlsxColumnReader() {
    }

    static List<String> read(Path file, int sheetNumber, int firstRow, String column,
                             int offset, Pattern separator) throws IOException {
        String baseColumn = column.strip().toUpperCase(Locale.ROOT);
        if (!baseColumn.matches("[A-Z]{1,3}")) {
            throw new IllegalArgumentException("Укажите букву столбца, например B или AA.");
        }
        int baseIndex = columnIndex(baseColumn);
        long targetIndex = (long) baseIndex + offset;
        if (baseIndex >= MAX_COLUMNS || targetIndex < 0 || targetIndex >= MAX_COLUMNS) {
            throw new IllegalArgumentException("Столбец со сдвигом выходит за пределы листа.");
        }
        if (firstRow < 1 || firstRow > MAX_ROWS) {
            throw new IllegalArgumentException("Некорректный номер первой строки: " + firstRow);
        }

        try (ZipFile zip = new ZipFile(file.toFile())) {
            DocumentBuilder builder = newDocumentBuilder();
            Element rootRelationships = readXml(zip, builder, "_rels/.rels");
            String workbookPath = relatedPart(rootRelationships, "", null, "officeDocument");
            Element workbook = readXml(zip, builder, workbookPath);
            List<Element> sheets = children(child(workbook, "sheets"), "sheet");
            if (sheetNumber < 1 || sheetNumber > sheets.size()) {
                throw new IllegalArgumentException("Лист " + sheetNumber
                        + " отсутствует. Листов в книге: " + sheets.size());
            }

            int slash = workbookPath.lastIndexOf('/') + 1;
            String relationshipsPath = workbookPath.substring(0, slash) + "_rels/"
                    + workbookPath.substring(slash) + ".rels";
            Element relationships = readXml(zip, builder, relationshipsPath);
            Element sheet = sheets.get(sheetNumber - 1);
            String sheetId = sheet.getAttributeNS(RELATIONSHIP_NS, "id");
            if (sheetId.isEmpty()) {
                sheetId = sheet.getAttributeNS(STRICT_RELATIONSHIP_NS, "id");
            }
            String sheetPath = relatedPart(relationships, workbookPath, sheetId, "worksheet");
            if (sheetPath == null) {
                throw new IOException("Выбранный лист не содержит таблицу с ячейками.");
            }

            String stringsPath = relatedPart(relationships, workbookPath, null, "sharedStrings");
            List<String> sharedStrings = stringsPath == null ? List.of()
                    : children(readXml(zip, builder, stringsPath), "si").stream()
                            .map(XlsxColumnReader::richText).toList();
            String stylesPath = relatedPart(relationships, workbookPath, null, "styles");
            List<Integer> zeroWidths = stylesPath == null ? List.of()
                    : zeroWidths(readXml(zip, builder, stylesPath));
            Element worksheet = readXml(zip, builder, sheetPath);
            return readRows(worksheet, (int) targetIndex, firstRow, separator, sharedStrings, zeroWidths);
        } catch (ParserConfigurationException | SAXException exception) {
            throw new IOException("Не удалось разобрать XML книги.", exception);
        }
    }

    private static List<String> readRows(Element worksheet, int targetIndex, int firstRow,
                                         Pattern separator, List<String> sharedStrings,
                                         List<Integer> zeroWidths) throws IOException {
        List<String> values = new ArrayList<>();
        int rowNumber = 0;
        for (Element row : children(child(worksheet, "sheetData"), "row")) {
            String rowReference = row.getAttribute("r");
            rowNumber = rowReference.isEmpty() ? rowNumber + 1 : Integer.parseInt(rowReference);
            if (rowNumber < firstRow) {
                continue;
            }

            int index = -1;
            for (Element cell : children(row, "c")) {
                String cellReference = cell.getAttribute("r");
                index = cellReference.isEmpty() ? index + 1 : columnIndex(cellReference);
                if (index == targetIndex) {
                    String type = cell.getAttribute("t");
                    String value = cellValue(cell, type, sharedStrings, zeroWidths);
                    boolean text = type.equals("s") || type.equals("inlineStr") || type.equals("str");
                    String[] parts = separator != null && text ? separator.split(value) : new String[]{value};
                    for (String part : parts) {
                        String trimmed = part.strip();
                        if (!trimmed.isEmpty()) {
                            values.add(trimmed);
                        }
                    }
                    break;
                }
            }
        }
        return values;
    }

    private static String cellValue(Element cell, String type, List<String> sharedStrings,
                                    List<Integer> zeroWidths) throws IOException {
        if (type.equals("inlineStr")) {
            return richText(child(cell, "is"));
        }
        String value = textContent(child(cell, "v"));
        if (value.isEmpty()) {
            return "";
        }

        return switch (type) {
            case "s" -> {
                int index = Integer.parseInt(value.strip());
                if (index < 0 || index >= sharedStrings.size()) {
                    throw new IOException("Некорректная ссылка на строку в ячейке " + cell.getAttribute("r"));
                }
                yield sharedStrings.get(index);
            }
            case "str" -> decodeExcelEscapes(value);
            case "b" -> "1".equals(value.strip()) || "true".equals(value.strip()) ? "TRUE" : "FALSE";
            case "", "n" -> formatNumber(cell, value.strip(), zeroWidths);
            default -> value;
        };
    }

    private static String formatNumber(Element cell, String value, List<Integer> zeroWidths) {
        String style = cell.getAttribute("s");
        int index = style.isEmpty() ? 0 : Integer.parseInt(style);
        int width = index >= 0 && index < zeroWidths.size() ? zeroWidths.get(index) : 0;
        if (width == 0 || value.isEmpty()) {
            return value;
        }
        BigDecimal number = new BigDecimal(value).setScale(0, RoundingMode.HALF_UP);
        String digits = number.abs().toPlainString();
        String sign = number.signum() < 0 ? "-" : "";
        return sign + "0".repeat(Math.max(0, width - digits.length())) + digits;
    }

    private static List<Integer> zeroWidths(Element styles) {
        Map<String, Integer> widths = new HashMap<>();
        widths.put("1", 1); // Встроенный числовой формат Excel "0".
        for (Element format : children(child(styles, "numFmts"), "numFmt")) {
            String code = format.getAttribute("formatCode");
            if (code.matches("0{1,255}")) {
                widths.put(format.getAttribute("numFmtId"), code.length());
            }
        }
        return children(child(styles, "cellXfs"), "xf").stream()
                .map(style -> widths.getOrDefault(style.getAttribute("numFmtId"), 0))
                .toList();
    }

    private static int columnIndex(String reference) {
        int index = 0;
        for (int i = 0; i < reference.length(); i++) {
            char letter = reference.charAt(i);
            if (letter < 'A' || letter > 'Z') {
                break;
            }
            index = index * 26 + letter - 'A' + 1;
        }
        return index - 1;
    }

    private static String richText(Element element) {
        StringBuilder value = new StringBuilder();
        for (Element part : children(element, null)) {
            switch (part.getLocalName()) {
                case "t" -> value.append(decodeExcelEscapes(textContent(part)));
                case "r" -> value.append(decodeExcelEscapes(textContent(child(part, "t"))));
                default -> { }
            }
        }
        return value.toString();
    }

    private static String decodeExcelEscapes(String value) {
        return EXCEL_ESCAPE.matcher(value).replaceAll(match -> {
            char character = (char) Integer.parseInt(match.group(1), 16);
            return Matcher.quoteReplacement(String.valueOf(character));
        });
    }

    private static String relatedPart(Element relationships, String source, String id, String type) {
        for (Element relationship : children(relationships, "Relationship")) {
            if (!relationship.getAttribute("Type").endsWith("/" + type)
                    || id != null && !id.equals(relationship.getAttribute("Id"))
                    || "External".equalsIgnoreCase(relationship.getAttribute("TargetMode"))) {
                continue;
            }
            URI target = URI.create("/" + source.replace(" ", "%20"))
                    .resolve(relationship.getAttribute("Target")).normalize();
            String path = target.getPath();
            if (target.getScheme() == null && target.getAuthority() == null
                    && path != null && path.startsWith("/") && !path.startsWith("/../")) {
                return path.substring(1);
            }
        }
        return null;
    }

    private static DocumentBuilder newDocumentBuilder() throws ParserConfigurationException {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newDefaultNSInstance();
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        return factory.newDocumentBuilder();
    }

    private static Element readXml(ZipFile zip, DocumentBuilder builder, String path)
            throws IOException, SAXException {
        var entry = path == null ? null : zip.getEntry(path);
        if (entry == null) {
            throw new IOException("В файле XLSX отсутствует обязательная часть"
                    + (path == null ? "." : " " + path + "."));
        }
        try (InputStream input = zip.getInputStream(entry)) {
            return builder.parse(input).getDocumentElement();
        }
    }

    private static Element child(Element parent, String name) {
        if (parent != null) {
            for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
                if (node instanceof Element element && name.equals(element.getLocalName())) {
                    return element;
                }
            }
        }
        return null;
    }

    private static List<Element> children(Element parent, String name) {
        List<Element> elements = new ArrayList<>();
        if (parent != null) {
            for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
                if (node instanceof Element element && (name == null || name.equals(element.getLocalName()))) {
                    elements.add(element);
                }
            }
        }
        return elements;
    }

    private static String textContent(Element element) {
        return element == null ? "" : element.getTextContent();
    }
}
