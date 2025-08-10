package task;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class Q4 {
    public static void main(String[] args) {
       
        Workbook workbook = new XSSFWorkbook();

        
        Sheet sheet = workbook.createSheet("Grocery List");

        
        String[][] groceries = {
                {"Item Name", "Quantity", "Price (₹)"},
                {"Rice", "5 Kg", "250"},
                {"Wheat Flour", "10 Kg", "450"},
                {"Sugar", "2 Kg", "90"},
                {"Milk", "2 Liters", "120"},
                {"Eggs", "12 Pieces", "70"}
        };

        
        for (int i = 0; i < groceries.length; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < groceries[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(groceries[i][j]);
            }
        }

        
        for (int i = 0; i < groceries[0].length; i++) {
            sheet.autoSizeColumn(i);
        }

        
        String filePath = ".//Data//GroceryList.xlsx";

        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
            System.out.println("Excel file created successfully at: " + filePath);
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

