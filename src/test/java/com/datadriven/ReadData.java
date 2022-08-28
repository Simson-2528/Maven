package com.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	//HOW TO READ DATAS FROM EXCEL
	
	
	public static void main(String[] args) throws IOException {

		
		File f = new File("C:\\Users\\Public\\Downloads\\Book1.xlsx");
		
		FileInputStream file = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(file);
		
		Sheet sheet = w.getSheet("Sheet1");
		
		
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			
			Row row = sheet.getRow(i);
			
			
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				
				Cell cell = row.getCell(j);
				
				DataFormatter d = new DataFormatter();
				
				String scv = d.formatCellValue(cell);
				System.out.println(scv);
				
				
				}
			
			
		}
		
		
		
		
		
	}

}
