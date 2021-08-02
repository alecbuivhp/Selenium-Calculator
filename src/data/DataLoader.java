/*
    Ref: https://viblo.asia/p/huong-dan-doc-va-ghi-file-excel-trong-java-su-dung-thu-vien-apache-poi-RQqKLENpZ7z
*/

package data;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataLoader {
    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_DESCRIPTION = 1;
    public static final int COLUMN_INDEX_NUM1 = 2;
    public static final int COLUMN_INDEX_NUM2 = 3;
    public static final int COLUMN_INDEX_OPERATOR = 4;
    public static final int COLUMN_INDEX_IS_INTEGER = 5;
    public static final int COLUMN_INDEX_EXPECTED_RESULT = 6;
    public static final int COLUMN_INDEX_DRIVER = 7;
    public static final int COLUMN_INDEX_BUILD = 8;

    public static Object[][] readExcel(String excelFilePath) throws IOException {
        List<Object[]> listTestCases = new ArrayList<>();

        // Get file
        InputStream inputStream = new FileInputStream(excelFilePath);

        // Get workbook
        Workbook workbook = getWorkbook(inputStream, excelFilePath);

        // Get sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Get all rows
        for (Row nextRow : sheet) {
            if (nextRow.getRowNum() == 0) {
                continue;
            }

            // Get all cells
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            // Read cells and set value for test case object
            Object[] testCase = new Object[9];

            while (cellIterator.hasNext()) {
                // Read cell
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                int columnIndex = cell.getColumnIndex();

                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                // Set value for test case object
                switch (columnIndex) {
                    case COLUMN_INDEX_ID:
                    case COLUMN_INDEX_OPERATOR:
                    case COLUMN_INDEX_DRIVER:
                    case COLUMN_INDEX_BUILD:
                        testCase[columnIndex] = BigDecimal.valueOf((double) cellValue).intValue();
                        break;
                    case COLUMN_INDEX_DESCRIPTION:
                        testCase[columnIndex] = (getCellValue(cell).toString());
                        break;
                    case COLUMN_INDEX_NUM1:
                    case COLUMN_INDEX_NUM2:
                    case COLUMN_INDEX_EXPECTED_RESULT:
                        if (cell.getCellType() == CellType.STRING) {
                            testCase[columnIndex] = (getCellValue(cell).toString());
                        } else {
                            String str = NumberToTextConverter.toText((double) cellValue);
                            testCase[columnIndex] = str;
                        }
                        break;
                    case COLUMN_INDEX_IS_INTEGER:
                        String boolValue = getCellValue(cell).toString();
                        testCase[columnIndex] = (boolValue.equalsIgnoreCase("true"));
                        break;
                    default:
                        break;
                }
            }
            if (testCase[0] != null) {
                listTestCases.add(testCase);
            }
        }

        workbook.close();
        inputStream.close();

        return listTestCases.toArray(new Object[listTestCases.size()][]);
    }

    // Get Workbook
    private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    // Get cell value
    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getStringValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            default:
                break;
        }

        return cellValue;
    }


}
