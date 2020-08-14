package utils;

import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTest {
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public ExcelTest() {}
	
	public ExcelTest(String excelPath) throws IOException {

			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheetAt(0);}
				
	public static Object getCellData(int rowNum,int colNum) throws IOException {		
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
		return value;

	}
	public static int getRowCount() {
		
		int rowCount = sheet.getPhysicalNumberOfRows();			
		return rowCount;
				
	}
	public static int getCellCount() {
		
		int rowCount = sheet.getPhysicalNumberOfRows();
		int maxCell = 0;
		for (int j = 0; j < rowCount; j++) {
			Row r = sheet.getRow(j);
			maxCell =  r.getLastCellNum();
		}
		return maxCell;
			
}
	
	


}
