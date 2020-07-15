package net.codejava.excel;
 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
 
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
 
/**
 * Create new excel sheet
 */

public class CreateExcelSheet {
 
 
    public static void main(String[] args) {
        String excelFilePath = "JavaBooks.xls";
         
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);
 
            Sheet newSheet = workbook.createSheet("Comments");
            Object[][] bookComments = {
                    {"Head First Java", "Funny and Exciting"},
                    {"Effective Java", "Insightful tips and advices"},
                    {"Clean Code", "Write Readable Code"},
                    {"Thinking in Java", "Classic"},
            };
      
            int rowCount = 0;
              
            for (Object[] aBook : bookComments) {
                Row row = newSheet.createRow(++rowCount);
                  
                int columnCount = 0;
                  
                for (Object field : aBook) {
                    Cell cell = row.createCell(++columnCount);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
                  
            }       
 
            FileOutputStream outputStream = new FileOutputStream("JavaBooks.xls");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
             
        } catch (IOException | EncryptedDocumentException
                | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }
 
}
