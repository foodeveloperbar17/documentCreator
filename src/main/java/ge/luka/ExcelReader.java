package ge.luka;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    public static String errorMessages = "";

    private XSSFSheet sheet;
    private int driverColumnIndex = -1;
    private int firstDayIndex;
    private int numDays;
    private int lastWorkingRowIndex = 2;

    private int firstNameIndex;
    private int lastNameIndex;
    private int carIdColumnIndex;
    private int timesIndex;
    private int destinationIndex;
    private int idIndex;

    private static final int HEADER_ROW_INDEX = 0;

    public ExcelReader(String filePath) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(filePath);
            sheet = workbook.getSheetAt(0);
        } catch (IOException e) {
            errorMessages += "ვერ ვიპოვე ფაილი \n";
            e.printStackTrace();
        }
    }

    public List<List<DocumentModel>> getAllTables() {
        errorMessages = "";
        boolean isSuccessful = initializeUtilityVariables();
        if (!isSuccessful) {
            return null;
        }
        List<List<DocumentModel>> result = new ArrayList<>();

        while (true) {
            List<DocumentModel> document = getDocument();
            if (lastWorkingRowIndex == -1) {
                break;
            }
            if (document.size() != 0) {
                result.add(document);
            }
        }

        return result;
    }

    public List<DocumentModel> getDocument() {
        List<DocumentModel> result = new ArrayList<>();
        int driverRowIndex = ExcelSearchUtils.getFirstNonNullStringRow(lastWorkingRowIndex, driverColumnIndex, sheet);
        if (driverRowIndex == -1) {
            errorMessages = "ბოლო მძღოლის სტრიქონის ნომერი: " + lastWorkingRowIndex + "\n" + errorMessages;
            lastWorkingRowIndex = -1;
            return result;
        }
        boolean isValidHours = checkValidHours(driverRowIndex);
        for (int i = 0; i < numDays; i++) {
            DocumentModel currDayDocument = new DocumentModel();
            setDriverName(currDayDocument, driverRowIndex);
            setClinicName(currDayDocument, driverRowIndex);
            if (isValidHours) {
                setTimes(currDayDocument, driverRowIndex);
            }
            setBlockId(currDayDocument, driverRowIndex);
            setClients(currDayDocument, driverRowIndex, firstDayIndex + i);
            setDay(currDayDocument, firstDayIndex + i);
            setCarId(currDayDocument, driverRowIndex);
            if (currDayDocument.getClients() != null && currDayDocument.getClients().size() != 0) {
                result.add(currDayDocument);
            }
        }
        lastWorkingRowIndex = driverRowIndex + 1;
        return result;
    }

    private void setBlockId(DocumentModel currDayDocument, int driverRowIndex) {
        XSSFCell cell = sheet.getRow(driverRowIndex).getCell(idIndex);
        if (cell == null) {
            currDayDocument.setBlockId("");
        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
            currDayDocument.setBlockId((int) cell.getNumericCellValue() + "");
        } else {
            currDayDocument.setBlockId(cell.toString());
        }
    }

    private boolean checkValidHours(int driverRowIndex) {
        XSSFCell cell = sheet.getRow(driverRowIndex).getCell(timesIndex);
        if (cell == null) {
            return false;
        }
        String times = cell.toString();
        String[] split = times.split(" +");
        if (split.length < 2) {
            errorMessages += "სფეისი არ აქვს დროებს. სტრიქონის მისამართი: " + (driverRowIndex + 1) + ". ტექსტი: " + times + "\n";
            return false;
        } else if (split.length > 2) {
            errorMessages += "ბევრი სფეისი აქვს დროებს. სტრიქონის მისამართი: " + (driverRowIndex + 1) + ". ტექსტი: " + times + "\n";
            return false;
        }
        return true;
    }

    private void setCarId(DocumentModel currDayDocument, int driverRowIndex) {
        XSSFCell cell = sheet.getRow(driverRowIndex).getCell(carIdColumnIndex);
        currDayDocument.setCarId(cell.toString());
    }

    private void setDay(DocumentModel currDayDocument, int dayColumnIndex) {
        XSSFCell dayCell = sheet.getRow(HEADER_ROW_INDEX).getCell(dayColumnIndex);
        currDayDocument.setDay(dayCell.toString());
    }

    private void setClinicName(DocumentModel result, int driverRowIndex) {
        XSSFCell cell = sheet.getRow(driverRowIndex).getCell(destinationIndex);
        result.setClinicName(cell.toString());
    }

    private void setClients(DocumentModel result, int driverRowIndex, int dayColumnIndex) {
        List<Client> clients = new ArrayList<>();
        int firstClientIndex = driverRowIndex + 1;
        int lastClientIndex = ExcelSearchUtils.getFirstNullRow(firstClientIndex, firstNameIndex, sheet) - 1;
        int numClients = lastClientIndex - firstClientIndex + 1;
        if (numClients == 0) {
            return;
        }

        for (int i = firstClientIndex; i <= lastClientIndex; i++) {
            Client client = getClientFromRow(i, dayColumnIndex);
            if (client != null) {
                clients.add(client);
            }
        }
        result.setClients(clients);
    }

    private Client getClientFromRow(int rowIndex, int dayColumnIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        XSSFCell isPresentColumn = row.getCell(dayColumnIndex);
        if (isPresentColumn != null && (isPresentColumn.toString().equals("1.0") || isPresentColumn.toString().equals("1"))) {
            Client client = new Client();
            client.setFirstName(row.getCell(firstNameIndex).toString());
            client.setLastName(row.getCell(lastNameIndex).toString());
            return client;
        } else {
            return null;
        }
    }

    private void setTimes(DocumentModel result, int driverRowIndex) {
        XSSFCell cell = sheet.getRow(driverRowIndex).getCell(timesIndex);
        String times = cell.toString();
        String[] split = times.split(" +");
        String start = split[0].trim();
        String end = split[1].trim();
        result.setStartHour(start);
        result.setEndHour(end);
    }

    private void setDriverName(DocumentModel documentModel, int driverRowIndex) {
        XSSFCell cell = sheet.getRow(driverRowIndex).getCell(driverColumnIndex);
        documentModel.setDriverFullName(cell.getStringCellValue());
    }


    private int getTextIndexInRow(XSSFRow row, String text) {
        if (text == null) {
            return -1;
        }
        for (int i = 0; i < 100; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null && cell.getCellType().equals(CellType.STRING) && text.equals(cell.getStringCellValue())) {
                return i;
            }
        }
        return -1;
    }

    public static char indexToChar(int index) {
        return (char) (index + 1 + 'a');
    }

    private boolean initializeUtilityVariables() {
        boolean success = true;
        XSSFRow headersRow = sheet.getRow(HEADER_ROW_INDEX);
        driverColumnIndex = getTextIndexInRow(headersRow, "მძღოლი");
        firstNameIndex = getTextIndexInRow(headersRow, "სახელი");
        lastNameIndex = getTextIndexInRow(headersRow, "გვარი");
        timesIndex = getTextIndexInRow(headersRow, "დრო");
        destinationIndex = getTextIndexInRow(headersRow, "კლინიკა");
        carIdColumnIndex = getTextIndexInRow(headersRow, "მანქანის ნომერი");
        idIndex = getTextIndexInRow(headersRow, "კოდი");

//        TODO: temp
//        firstNameIndex = 2;
//        lastNameIndex = 3;
//        timesIndex = 8;
//        destinationIndex = 9;
//        carIdColumnIndex = 5;
//        idIndex = 12;


        if (driverColumnIndex == -1) {
            errorMessages += "სვეტი მძღოლი ვერ მოიძებნა\n";
            success = false;
        }
        if (firstNameIndex == -1) {
            errorMessages += "სვეტი სახელი ვერ მოიძებნა\n";
            success = false;
        }
        if (lastNameIndex == -1) {
            errorMessages += "სვეტი გვარი ვერ მოიძებნა\n";
            success = false;
        }
        if (timesIndex == -1) {
            errorMessages += "სვეტი დრო ვერ მოიძებნა\n";
            success = false;
        }
        if (destinationIndex == -1) {
            errorMessages += "სვეტი კლინიკა ვერ მოიძებნა\n";
            success = false;
        }
        if (carIdColumnIndex == -1) {
            errorMessages += "სვეტი მანქანის ნომერი ვერ მოიძებნა\n";
            success = false;
        }
        if (idIndex == -1) {
            errorMessages += "სვეტი კოდი ვერ მოიძებნა";
            success = false;
        }

        firstDayIndex = ExcelSearchUtils.firstNonNullColumnIndex(headersRow, driverColumnIndex + 1);
        if (firstDayIndex == -1) {
            errorMessages += "დღეების სათაურები ვერ მოიძებნა";
            success = false;
        }
        int lastDayIndex = ExcelSearchUtils.firstNullColumnIndex(headersRow, firstDayIndex) - 1;
        numDays = lastDayIndex - firstDayIndex + 1;
        if (numDays <= 0) {
            errorMessages += "შეცდომა დღეებში";
            success = false;
        }
        return success;
    }
}
