package utils;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;

public class ExcelMain {
	
	@SuppressWarnings("static-access")
	public static void main(String[] args) throws IOException {
		
		@SuppressWarnings("resource")
		Scanner sc = new Scanner(System.in);
		String excelPath = "";
		ExcelTest excel = new ExcelTest();
		boolean exist = true;
		do {
		try {
		exist = true;
		System.out.println("Zadej cestu k Excel souboru (.xlsx)\nNápověda: například C:\\Users\\Desktop\\vzorek_Dat.xlsx");
		excelPath = sc.nextLine();
		excel = new ExcelTest(excelPath);
		}
		
		catch(FileNotFoundException exp) {System.out.println("Špatně zadaná cesta!");exist=false;}
		catch(InvalidOperationException exp) {System.out.println("Špatně zadaný formát cesty!");exist=false;}
		catch(IllegalArgumentException exp) {System.out.println("Špatně zadáno, zadali jste cestu složky!");exist=false;}
		
		}while(exist != true); 
		
		System.out.printf("Prvočísla ve Vámi vybraném EXCEL souboru");
				
		excel.prvocisla();	

	}
}
