package org.automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Excel {
	
	@Test
	public void getExce() throws IOException
	{
//		Object[][] data = { {}, {}, {}};
		
		//every row of excel should be sent to 1 arry 
		File loc=new File("C:\\Users\\acer\\Downloads\\datadrivenTestngDemo1.xlsx"); 
		FileInputStream fis = new FileInputStream(loc);
		XSSFWorkbook wbook = new XSSFWorkbook(fis); 
		XSSFSheet sheet = wbook.getSheetAt(0);
		int rowcount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int columncount = row.getLastCellNum();
		
		Object[][] data = new Object[rowcount-1] [columncount];
//		data[0][0]=hello
//		data[0][1]=text
//		data[0][2]=id
		
		for(int i=0; i<rowcount-1; i++)
		{
			System.out.println("outer loop started =========>");
			row=sheet.getRow(i+1);
			for(int j=0; j<columncount; j++)
					{
				XSSFCell cell = row.getCell(j);
				
				System.out.println(cell);
						
					}
			System.out.println("outerloop ended ==========>");
		}
	}

}
