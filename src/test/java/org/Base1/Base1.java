package org.Base1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Base1 {
	
	public static void write() throws IOException {
		String path = "C:\\Users\\Admin\\eclipse-workspace\\TestNg4\\src\\data\\JMSA.xlsx";
		FileInputStream fis = new FileInputStream(path);
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();for(
		int i = 0;i<=lastRow;i++)
		{
			Row row = sheet.getRow(0);
			Row row2 = sheet.getRow(1);
			Cell cell = row.createCell(9);
			cell.setCellValue("status");

		}
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();

	
	
		
		
		

	}

}
