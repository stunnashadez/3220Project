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
        String excelFilePath = "CommunityCentre.xlsx";
         
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);
 
            Sheet newSheet = workbook.createSheet("Directions");
            Object[][] bookComments = {
                    {"Little River Golf Course", "Near Tecumseh Mall"},
                    {"Roseland Golf and Curling Club", "Near Roseland"},
                    {"Forest Glade Community Centre", "Near Forest Glade"},
                    {"Ojibway Nature Centre", "Near Ojibway Park"},
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
 
            FileOutputStream outputStream = new FileOutputStream("CommunityCentre.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
             
        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }
 
}
