package com.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ddf.EscherColorRef.SysIndexSource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

							//HOW TO WRITE DATAS IN EXCEL

public class WriteData {

	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\Public\\Downloads\\Book1.xlsx");
		
		FileInputStream file = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(file);
		
		Sheet cs = w.createSheet("Sheet1");
		//Applying value to excel
		
		Object a[][]= {{"Simson", 7904213337l, 16001, "Gunasekaran", "Suriya", "chennai", "DECE", 2019, 79.8f},
						{"pisasu", 6543516413l, 16002, "mayilu", "bujjuks", "ooty", "Bd", 2019, 71.8f},
						{"Guna", 6599413l, 16002, "simson", "shantha", "ramapuram", "BE", 2019, 77.8f},
						{"santhosh", 99456413l, 16002, "subbu", "paichaiyamma", "CBE", "BA", 2019, 77.5f},
						{"pavithra", 654112413l, 16002, "parasu", "priyanka", "itali", "Bsc", 2019, 75.8f},
						{"saara", 6543516545l, 16002, "kuttyboy", "gundu", "ramapuram", "Btech", 2019, 87.8f}};
	

		
		//create row
		for (int i = 0; i < a.length; i++) {
			
			Row cr = cs.createRow(i);
			
			
			//create cells
			
			for (int j = 0; j < a[i].length; j++) {
				
				Cell cc = cr.createCell(j);
				
			//Object for array 	
				
				Object value = a[i][j];
				
				//conditions for given values
				
				if (value instanceof String) {
					
					cc.setCellValue((String)value);
				} 
				
				
				else if(value instanceof Boolean){
					cc.setCellValue((Boolean)value);
					
					
					}
				else if(value instanceof Integer){
					cc.setCellValue((Integer)value);
					
					
					}
				
				else if(value instanceof Float){
					cc.setCellValue((Float)value);
					}
				
				else if(value instanceof Long){
					cc.setCellValue((Long)value);
					
				}
			
				
			}
			
		}
		
			FileOutputStream fo = new FileOutputStream(f);
			
			w.write(fo);
				
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
