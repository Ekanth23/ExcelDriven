package org.automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataDrivenTestNg {
	
	DataFormatter formatter=new DataFormatter();
	
	@Test(dataProvider="driveTest")
	public void testCaseData(String greeting, String communication, String id)
	{
		System.out.println(greeting +" : "+ communication +" : "+ id);
	}
	
	@DataProvider(name="driveTest")
	public Object[][] getData() throws IOException
	{
		
		
		//every row of excel should be sent to 1 arry 
		File loc=new File("C:\\Users\\acer\\Downloads\\datadrivenTestngDemo1.xlsx"); 
		FileInputStream fis = new FileInputStream(loc);
		XSSFWorkbook wbook = new XSSFWorkbook(fis); 
		XSSFSheet sheet = wbook.getSheetAt(0);
		int rowcount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int columncount = row.getLastCellNum();
//		int columncount =3; 
		
		Object[][] data = new Object[rowcount-1] [columncount];

		for(int i=0; i<rowcount-1; i++)
		{
			row=sheet.getRow(i+1);
			for(int j=0; j<columncount; j++)
					{
				XSSFCell cell = row.getCell(j);
				data[i][j]=formatter.formatCellValue(cell);
		
					}
		}
		
		return data;
	}

}
