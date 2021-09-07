package ge.luka;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelWriter {

    int rowIndex = 0;

    private static final int HEADER_COLUMN_INDEX = 5;
    private static final String DATE_SUFFIX = ".2021";
    private XSSFWorkbook xssfWorkbook;

    public void writeDocument(List<List<DocumentModel>> allTables, String folderPath) {
        if (allTables == null) {
            return;
        }
        for (List<DocumentModel> allTable : allTables) {
            if (allTable.size() == 0) {
                continue;
            }
            rowIndex = 0;
            xssfWorkbook = new XSSFWorkbook();
            XSSFSheet sheet = xssfWorkbook.createSheet("Sheet 1");
            formatSheet(sheet);
            for (DocumentModel documentModel : allTable) {
                writeOneDay(sheet, documentModel);
            }
            DocumentModel documentModel = allTable.get(0);
            String path = documentModel.getClinicName() + " " + documentModel.getStartHour() + " " + documentModel.getEndHour()
                    + " " + documentModel.getBlockId() + " " + documentModel.getDriverFullName();
            path = path.replaceAll("[^ ა-ჰA-Za-z0-9()\\[\\]]", "");
            path = folderPath + path + ".xlsx";
            saveWorkbook(xssfWorkbook, path);
        }
    }

    private void formatSheet(XSSFSheet sheet) {
        sheet.setColumnWidth(0, (int) (350 * 1.8));
        sheet.setColumnWidth(1, 350 * 12);
        sheet.setColumnWidth(2, (int) (350 * 19.5));
        sheet.setColumnWidth(3, (int) (350 * 11.5));
        sheet.setColumnWidth(4, (int) (350 * 11.5));
        sheet.setColumnWidth(5, (int) (350 * 8.5));
        sheet.setDisplayGridlines(false);
    }

    private void writeOneDay(XSSFSheet sheet, DocumentModel documentModel) {
        rowIndex++;
        addHeader(sheet);
        addClinicName(sheet, documentModel);
        addCarId(sheet, documentModel);
        addDate(sheet, documentModel);
        addHeaders(sheet);
        addClients(sheet, documentModel);
        rowIndex += 24 - documentModel.getClients().size();
        addDriverName(sheet, documentModel);
        rowIndex += 2;
        addSignerCell(sheet, documentModel);
        rowIndex += 6;
        addClinicPersonName(sheet);
    }

    private void addSignerCell(XSSFSheet sheet, DocumentModel documentModel) {
        XSSFRow row = sheet.createRow(rowIndex);
        XSSFCell bottomBorderCell = row.createCell(3);
        XSSFCellStyle borderStyle = xssfWorkbook.createCellStyle();
        borderStyle.setBorderBottom(BorderStyle.THICK);
        bottomBorderCell.setCellStyle(borderStyle);
        rowIndex++;
    }

    private void addClinicPersonName(XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(rowIndex);

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 2));
        XSSFCell cell = row.createCell(1);
        cell.setCellValue("კლინიკის უფლება მოსილი პირის ხელმოწერა ");

        XSSFCellStyle documentHeaderStyle = getDocumentHeaderStyle();
        cell.setCellStyle(documentHeaderStyle);

        XSSFCell bottomBorderCell = row.createCell(3);
        XSSFCellStyle borderStyle = xssfWorkbook.createCellStyle();
        borderStyle.setBorderBottom(BorderStyle.THICK);
        bottomBorderCell.setCellStyle(borderStyle);

        rowIndex++;
    }

    private void addDriverName(XSSFSheet sheet, DocumentModel documentModel) {
        XSSFRow row = sheet.createRow(rowIndex);

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 2));
        XSSFCell cell = row.createCell(1);
        cell.setCellValue("ტრანსპორტირებაზე პასუხისმგებელი პირი ");

        XSSFCellStyle documentHeaderStyle = getDocumentHeaderStyle();
        cell.setCellStyle(documentHeaderStyle);

        XSSFCell bottomBorderCell = row.createCell(3);
        XSSFCellStyle borderStyle = xssfWorkbook.createCellStyle();
        borderStyle.setBorderBottom(BorderStyle.THICK);
        bottomBorderCell.setCellStyle(borderStyle);
        bottomBorderCell.setCellValue(documentModel.getDriverFullName());

        rowIndex++;
    }

    private void addClients(XSSFSheet sheet, DocumentModel documentModel) {
        for (int i = 1; i <= documentModel.getClients().size(); i++) {
            Client client = documentModel.getClients().get(i - 1);
            XSSFRow row = sheet.createRow(rowIndex);
            XSSFCellStyle clientCellStyle = getClientCellStyle();

            XSSFCell indexCell = row.createCell(0);
            indexCell.setCellValue(i);
            indexCell.setCellStyle(clientCellStyle);

            XSSFCell nameCell = row.createCell(1);
            nameCell.setCellValue(client.getFirstName());
            nameCell.setCellStyle(clientCellStyle);

            XSSFCell lastNameCell = row.createCell(2);
            lastNameCell.setCellValue(client.getLastName());
            lastNameCell.setCellStyle(clientCellStyle);

            XSSFCell startTimeCell = row.createCell(3);
            startTimeCell.setCellValue(documentModel.getStartHour());
            startTimeCell.setCellStyle(clientCellStyle);

            XSSFCell endTimeCell = row.createCell(4);
            endTimeCell.setCellValue(documentModel.getEndHour());
            endTimeCell.setCellStyle(clientCellStyle);

            XSSFCell infoCell = row.createCell(5);
            infoCell.setCellValue("");
            infoCell.setCellStyle(clientCellStyle);

            rowIndex++;
        }
    }

    private XSSFCellStyle getClientCellStyle() {
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();

        XSSFFont font = xssfWorkbook.createFont();
        font.setBold(false);
        cellStyle.setFont(font);

        setAllBorders(cellStyle);
        return cellStyle;
    }

    private void addHeaders(XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(rowIndex);

        row.setHeightInPoints(100);
        XSSFCellStyle headerStyle = getHeaderStyle();

        XSSFCell nCell = row.createCell(0);
        nCell.setCellValue("N");
        nCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 2));
        XSSFCell fullNameCell = row.createCell(1);
        fullNameCell.setCellValue("ჰემოდიალიზის კომპონენტით მოსარგებლე  პაციენტები (სახელი გვარი)");
        XSSFCell secondFullNameCell = row.createCell(2);
        secondFullNameCell.setCellStyle(headerStyle);
        fullNameCell.setCellStyle(headerStyle);

        XSSFCell startTimeCell = row.createCell(3);
        startTimeCell.setCellValue("ჰემოდიალიზით კომპონენტით მოსარგებლე  პაციენტების  პროცედურის დაწყების  დრო");
        startTimeCell.setCellStyle(headerStyle);

        XSSFCell endTimeCell = row.createCell(4);
        endTimeCell.setCellValue("ჰემოდიალიზით კომპონენტით მოსარგებლე  პაციენტების  პროცედურის დასრულების  დრო");
        endTimeCell.setCellStyle(headerStyle);

        XSSFCell infoCell = row.createCell(5);
        infoCell.setCellValue("შენიშვნა");
        infoCell.setCellStyle(headerStyle);

        rowIndex++;
    }

    private XSSFCellStyle getHeaderStyle() {
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFFont font = xssfWorkbook.createFont();
        font.setBold(false);
        cellStyle.setFont(font);

        setAllBorders(cellStyle);
        return cellStyle;
    }

    private void setAllBorders(XSSFCellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
    }


    private void addDate(XSSFSheet sheet, DocumentModel documentModel) {
        XSSFRow row = sheet.createRow(rowIndex);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 6));
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("პროცედურის შესრულების თარიღი " + documentModel.getDay() + DATE_SUFFIX);

        XSSFCellStyle documentHeaderStyle = getDocumentHeaderStyle();
        cell.setCellStyle(documentHeaderStyle);

        rowIndex++;
    }

    private void addCarId(XSSFSheet sheet, DocumentModel documentModel) {
        XSSFRow row = sheet.createRow(rowIndex);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 6));
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("ავტოტრანსპორტი სახ. N " + documentModel.getCarId());

        XSSFCellStyle documentHeaderStyle = getDocumentHeaderStyle();
        cell.setCellStyle(documentHeaderStyle);

        rowIndex++;
    }

    private void addHeader(XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(rowIndex++);
        XSSFCell cell = row.createCell(HEADER_COLUMN_INDEX);
        cell.setCellValue("დანართი N2");

        XSSFCellStyle documentHeaderStyle = getDocumentHeaderStyle();
        cell.setCellStyle(documentHeaderStyle);

    }

    private XSSFCellStyle getDocumentHeaderStyle() {
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        XSSFFont font = xssfWorkbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private void addClinicName(XSSFSheet sheet, DocumentModel documentModel) {
        XSSFRow row = sheet.createRow(rowIndex);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 6));
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("კლინიკის დასახელება " + documentModel.getClinicName());

        XSSFCellStyle documentHeaderStyle = getDocumentHeaderStyle();
        cell.setCellStyle(documentHeaderStyle);

        rowIndex++;
    }

    private void saveWorkbook(XSSFWorkbook xssfWorkbook, String path) {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(path);
            xssfWorkbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
            ExcelReader.errorMessages += "ვერ ჩავწერე ფაილში. სავარაუდოდ ფაილი გახსნილია \n";
        }
    }
}
