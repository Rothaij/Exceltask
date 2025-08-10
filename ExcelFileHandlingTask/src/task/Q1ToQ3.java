package task;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class Q1ToQ3 {
	
    public static void main(String[] args) {
        
        Workbook workbook = new XSSFWorkbook();

        
        Sheet sheet = workbook.createSheet("SHEET1");

        
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Age");
        headerRow.createCell(2).setCellValue("Email");

        // Data to insert
        Object[][] data = {
            {"John Doe", 30, "john@test.com"},
            {"Jane Doe", 28, "john@test.com"},
            {"Bob Smith", 35, "jacky@example.com"},
            {"Swapnil", 37, "swapnil@example.com"}
        };

        
        int rowNum = 1;
        for (Object[] rowData : data) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object value : rowData) {
                Cell cell = row.createCell(colNum++);
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                }
            }
        }

        
        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        
        String excelFilePath =".\\Data\\Workbook.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
            System.out.println("Excel file created: WorkbookWithData.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }

       
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
