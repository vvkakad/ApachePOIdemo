package com.apachePOI;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WriteEx1 {
	
	public static org.apache.poi.ss.usermodel.Sheet sh = null;
	public static Row row = null;
	public static Cell cell = null;

	public static void main(String[] args) throws Exception {
		
		FileInputStream fis = new FileInputStream("MyFirstTestCase.xlsx");
		Workbook wb = new WorkbookFactory().create(fis);
		
		if(wb.getSheet("vvk")== null)
			sh =wb.createSheet("vvk");
		else
			sh= wb.getSheet("vvk");
		
		
		if(sh.getRow(4)==null)
			row = sh.createRow(4);
		else
			row = sh.getRow(4);
		
		
		if(row.getCell(3)==null)
			cell = row.createCell(3);
		else
			cell = row.getCell(3);
		
		cell.setCellValue("vvkakad");
		FileOutputStream fos = new FileOutputStream("MyFirstTestCase.xlsx");
		
		wb.write(fos);
		wb.close();
		fos.close();
		
	}
}





