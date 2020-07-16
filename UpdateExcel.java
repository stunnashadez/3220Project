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
 * Append new rows to an existing sheet.
 */

public class UpdateExcel 
{
     public static void main(String[] args) 
     {
        String excelFilePath = "CommunityCentre.xlsx";
         
        try 
        {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);
 
            Sheet sheet = workbook.getSheetAt(0);
 
            Object[][] bookData = 
            {
                    {"8787 McHugh St", "WFCU Centre", -82.92748649, 42.31871719},
            };
 
            int rowCount = sheet.getLastRowNum();
 
            for (Object[] aBook : bookData) 
            {
                Row row = sheet.createRow(++rowCount);
 
                int columnCount = 0;
                 
                Cell cell = row.createCell(columnCount);
                cell.setCellValue(rowCount);
                 
                for (Object field : aBook) 
                {
                    cell = row.createCell(++columnCount);
                    if (field instanceof String) 
                    {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) 
                    {
                        cell.setCellValue((Integer) field);
                    }
                }
            }
 
            inputStream.close();
 
            FileOutputStream outputStream = new FileOutputStream("CommunityCentre.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
             
        }
        catch (IOException | EncryptedDocumentException | InvalidFormatException ex) 
        {
            ex.printStackTrace();
        }
    }
}
