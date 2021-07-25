package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		String excelFilePath = ".\\datafiles\\Capital_Cities.xlsx";
		FileInputStream inputstream = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
				//XSSFSheet sheet = workbook.getSheet("Data");
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		
		for (int r=0; r<=rows; r++) {
			XSSFRow row = sheet.getRow(r);
			for (int c=0; c<cols; c++ ) {
				XSSFCell cell = row.getCell(c);
				
				
				
				switch (cell.getCellType()) {
				case STRING: System.out.println(cell.getStringCellValue());break;
				case NUMERIC: System.out.println(cell.getNumericCellValue());break;
				case BOOLEAN: System.out.println(cell.getBooleanCellValue());break;
				default: break;
				
				}
				
			}
		}
		System.out.println(1);

	}

}
