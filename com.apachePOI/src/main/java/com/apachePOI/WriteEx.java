package com.apachePOI;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WriteEx {
	
	public static org.apache.poi.ss.usermodel.Sheet sh = null;
	public static Row row = null;
	public static Cell cell = null;
	
	public static void main(String[] args) throws Exception {
		
		FileInputStream fis = new FileInputStream("MyFirstTestCase.xlsx");
		Workbook wb= WorkbookFactory.create(fis);
		
		if(wb.getSheet("vaibhav")==null)
			sh= wb.createSheet("vaibhav");
		
		else
			sh= wb.getSheet("vaibhav");
		
		if(sh.getRow(5)==null)
			row= sh.createRow(5);
		
		else
			row= sh.getRow(5);
		
		if(row.getCell(3)==null)
			cell= row.createCell(3);
		else
			cell= row.getCell(3);
		
		cell.setCellValue("MyFirstTestCase");
		FileOutputStream fos = new FileOutputStream("MyFirstTestCase.xlsx");
		
		wb.write(fos);
		wb.close();
		fos.close();
	}
}
