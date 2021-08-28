package ge.luka;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ExcelSearchUtils {

    public static int firstNonNullColumnIndex(XSSFRow row, int startIndex) {
        for (int i = startIndex; i < 100; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null) {
                return i;
            }
        }
        throw new RuntimeException("Couldn't find non-null column");
    }

    public static int firstNullColumnIndex(XSSFRow row, int startIndex) {
        for (int i = startIndex; i < 100; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell == null) {
                return i;
            }
        }
        throw new RuntimeException("Couldn't find null column");
    }

    public static boolean isCellEmpty(XSSFCell cell) {
        return cell == null || cell.toString().equals("") || cell.toString().equals("null");
    }

    public static int getFirstNullRow(int startRowIndex, int columnIndex, XSSFSheet sheet) {
        for (int i = startRowIndex; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(columnIndex);
            if (cell == null || cell.toString().equals("")) {
                return i;
            }
        }
        System.out.println("No more null row in this column columnIndex: " + columnIndex + " rowIndex: " + startRowIndex);
        return -1;
    }

    public static int getFirstNonNullStringRow(int startRowIndex, int columnIndex, XSSFSheet sheet) {
        for (int i = startRowIndex; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(columnIndex);
            if (cell != null && !cell.toString().equals("") && cell.getCellType().equals(CellType.STRING)
                    && !cell.toString().trim().equals("")) {
                return i;
            }
        }
        System.out.println("No more non-null rows in this column columnIndex: " + columnIndex + " rowIndex: " + startRowIndex);
        return -1;
    }

}
