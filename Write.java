import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class Write {
	public static void main(String[] args) {
		
		Write excel = new Write();
		excel.write();
	}
	
	public void write() {
		//Set the Stream to connect to excel file - To open the file
		try {
			FileInputStream fis = new FileInputStream ("C:\\Users\\joesa\\OneDrive\\Desktop\\ReadandWrite\\Book2.xlsx");
			//Open the Workbook
			XSSFWorkbook xlWorkbook = new XSSFWorkbook(fis);
			//Open the Sheet
			XSSFSheet xlSheet = xlWorkbook.getSheetAt(0);
			//Get hold of the Rows in the particular
			XSSFRow xlRow = xlSheet.createRow(0);
			//Now use Cells and write your data in to the cell
			XSSFCell xlCell = xlRow.createCell(0);
			//Row number 1 is created and updated with values
			xlCell.setCellValue("Name");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("Age");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("Email");
			//Row number 2 is created and updated with values
			xlRow = xlSheet.createRow(1);
			xlCell = xlRow.createCell(0);
			xlCell.setCellValue("John Doe");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("30");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("john@test.com");
			//Row number 3 is created and updated with values
			xlRow = xlSheet.createRow(2);
			xlCell = xlRow.createCell(0);
			xlCell.setCellValue("Jane Doe");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("28");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("john@test.com");
			xlRow = xlSheet.createRow(3);
			xlCell = xlRow.createCell(0);
			xlCell.setCellValue("Bob Smith");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("35");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("jacky@example.com");
			xlRow = xlSheet.createRow(4);
			xlCell = xlRow.createCell(0);
			xlCell.setCellValue("Swapnil");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("37");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("swapnil@example.com");
			
			//OutStream to write the values to the destination file
			FileOutputStream fos = new FileOutputStream("C:\\Users\\joesa\\OneDrive\\Desktop\\ReadandWrite\\Book2.xlsx");
			xlWorkbook.write(fos);
			fis.close();
			fos.close();
			xlWorkbook.close();
			
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}