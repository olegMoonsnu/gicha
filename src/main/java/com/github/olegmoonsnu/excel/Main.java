package com.github.olegmoonsnu.excel;

import com.google.gson.GsonBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;

public class Main {

    // Настройки. Номера листов и строк начинаются с 1, как в Excel.
    private static final String EXCEL_FILE = "input.xlsx";
    private static final int SHEET_NUMBER = 1;
    private static final int FIRST_DATA_ROW = 2;

    private static final String COLUMN = "B";
    private static final int COLUMN_OFFSET = 0; // -1: A, 0: B, +1: C

    private static final boolean SPLIT_BY_COMMA = true;
    private static final boolean SPLIT_BY_NEWLINE = true;

    public static void main(String[] args) {
        if (args.length > 1) {
            System.err.println("Передайте один аргумент: путь к Excel-файлу.");
            System.exit(1);
        }

        try {
            Path file = Path.of(args.length == 1 ? args[0] : EXCEL_FILE);
            List<String> values = readColumn(file);
            String json = new GsonBuilder().disableHtmlEscaping().create().toJson(values);
            System.out.println(json);
        } catch (IOException | RuntimeException exception) {
            System.err.println("Ошибка чтения Excel: " + exception.getMessage());
            System.exit(1);
        }
    }

    private static List<String> readColumn(Path file) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(file.toFile(), null, true)) {
            if (SHEET_NUMBER < 1 || SHEET_NUMBER > workbook.getNumberOfSheets()) {
                throw new IllegalArgumentException("Лист " + SHEET_NUMBER
                        + " отсутствует. Листов в книге: " + workbook.getNumberOfSheets());
            }

            if (FIRST_DATA_ROW < 1 || FIRST_DATA_ROW > workbook.getSpreadsheetVersion().getMaxRows()) {
                throw new IllegalArgumentException("Некорректный номер первой строки: " + FIRST_DATA_ROW);
            }

            int columnIndex = columnIndex(workbook);
            Sheet sheet = workbook.getSheetAt(SHEET_NUMBER - 1);
            DataFormatter formatter = new DataFormatter(Locale.ROOT);
            formatter.setUseCachedValuesForFormulaCells(true);
            Pattern separator = createSeparator();

            List<String> values = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() + 1 < FIRST_DATA_ROW) {
                    continue;
                }

                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null) {
                    continue;
                }

                CellType type = cell.getCellType();
                if (type == CellType.FORMULA) {
                    type = cell.getCachedFormulaResultType();
                }

                String text = formatter.formatCellValue(cell);
                String[] parts = separator != null && type == CellType.STRING
                        ? separator.split(text)
                        : new String[]{text};
                for (String part : parts) {
                    String value = part.strip();
                    if (!value.isEmpty()) {
                        values.add(value);
                    }
                }
            }
            return values;
        }
    }

    private static int columnIndex(Workbook workbook) {
        String column = COLUMN.strip().toUpperCase(Locale.ROOT);
        if (!column.matches("[A-Z]{1,3}")) {
            throw new IllegalArgumentException("Укажите букву столбца, например B или AA.");
        }

        int baseColumnIndex = CellReference.convertColStringToIndex(column);
        long index = (long) baseColumnIndex + COLUMN_OFFSET;
        int maxColumns = workbook.getSpreadsheetVersion().getMaxColumns();
        if (baseColumnIndex >= maxColumns || index < 0 || index >= maxColumns) {
            throw new IllegalArgumentException("Столбец " + column + " со сдвигом "
                    + COLUMN_OFFSET + " выходит за пределы формата книги.");
        }
        return (int) index;
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
