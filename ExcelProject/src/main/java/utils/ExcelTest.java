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
	public static void prvocisla() throws IOException {
		int maxCEL = getCellCount();
		int maxROW = getRowCount();
		int prvocislo = 0;
		String test = "";
		
		for (int i = 0; i < maxROW; i++) {
			for (int j = 0; j < maxCEL; j++) {
				if ((getCellData(i, j)) != "" && (getCellData(i, j)) != null) {
				test = (getCellData(i, j)).toString();
			    try {
			    	prvocislo = Integer.parseInt(test);
			    } 
			    catch (NumberFormatException e) {}
			    
			    if (((prvocislo > 1) && (prvocislo % 2 != 0) && (prvocislo % 3 != 0) && (prvocislo % 5 != 0)) || prvocislo == 2 || prvocislo == 3 ) { 
			    	System.out.printf("%d\n", prvocislo);}
			    }
			}			
		}		
	}	

}
