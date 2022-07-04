package com.apachePOI;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadEx {
	public static void main(String[] args)throws Exception {
		DataFormatter df= new DataFormatter();
		FileInputStream fis = new FileInputStream("MySecondTestCase.xlsx");
		Workbook wb= WorkbookFactory.create(fis);
		
		Sheet sh= wb.getSheet("Test Cases");
		
		int rows = sh.getLastRowNum();
		
		for(int i=0;i<=rows;i++)
		{
			int col=sh.getRow(i).getLastCellNum();
			for(int j=0;j<col;j++)
			{
				Cell c= sh.getRow(i).getCell(j);
				System.out.print(df.formatCellValue(c)+"  ");
			}
			System.out.println();
		}
	
	}
}
