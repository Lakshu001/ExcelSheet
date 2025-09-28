**1**
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelWorkbook1 {
    public static void main(String[] args) {
    	try (Workbook workbook = new XSSFWorkbook();
    			FileOutputStream fos = new FileOutputStream  ("NewWorkbook.xlsx")) {
    				
    				
    				workbook.createSheet ("Sheet1");
    				workbook.write (fos);
    				System.out.println("New Excel Workbook Created : NewWorkbook.xls");
    	}    	catch (Exception e){
    			e.printStackTrace();
    }
  }  
}

**2**
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sheet1 {
	
	public static void main (String[] args) {
		
		Workbook workbook = new XSSFWorkbook ();
		
		workbook.createSheet ("Sheet 1");
		System.out.println("Workbook 'Sheet1' created Successfully");
				
	}
}

**3**
import java.io.FileWriter;
import java.io.IOException;

public class DataDetials {

	public static void main (String[] args){ 
		
		String[] titles = {"Name", "Age", "Email"};
		String[][] data = {
				
				{"john doe", "30", "john@test.com"},
				{"john doe", "28", "john@test.com"},
				{ "Bob Smith", "35", "jacky@example.com" },
	            { "Swapnil", "37", "swapnil@example.com"}
		};	
			
		try(FileWriter writter = new FileWriter("output.csv")){
			
			writter.append(String.join(",", titles)).append("\n");
			
			for (String[] row : data) {
				writter.append(String.join(",", row)).append("\n");
				}
		System.out.println("CSV file written successfully.");
		}catch (IOException e) {
            e.printStackTrace();
		}
	}
}

**4**

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriiter {

    public static void main(String[] args) {
        String[] headers = { "Name", "Age", "Email" };
        String[][] data = {
            { "John Doe", "30", "john@test.com" },
            { "Jane Doe", "28", "john@test.com" },
            { "Bob Smith", "35", "jacky@example.com" },
            { "Swapnil", "37", "swapnil@example.com" }
        };
 
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("User Data");
        
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                row.createCell(j).setCellValue(data[i][j]);
            }
        }
        
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream out = new FileOutputStream("output.xlsx")) {
            workbook.write(out);
            workbook.close();
            System.out.println("Excel file written successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


**5**
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	public static void main(String[] args) {
        String excelFilePath = "output.xlsx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                   
            switch (cell.getCellType()) {
            case STRING:
                System.out.print(cell.getStringCellValue() + "\t");
                       break;
                      case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                      System.out.print(cell.getDateCellValue() + "\t");
                     } else {
                System.out.print((int) cell.getNumericCellValue() + "\t");                        }
                          break;
                      case BOOLEAN:
                          System.out.print(cell.getBooleanCellValue() + "\t");
                          break;
                      case FORMULA:
                          System.out.print(cell.getCellFormula() + "\t");
                          break;
                      default:
                          System.out.print(" \t");
                  }
                }
                System.out.println(); 
                }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
