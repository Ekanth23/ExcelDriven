package org.automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	// identify test case column header
	// once column header is identified thenn scan entire testcase coulnm to
	// identify puchse testcase row
	// after you grab purchase testcase for = pull all the data of that row and then
	// feed it to the test case

	public ArrayList<String> getData(String testcasename) throws IOException {
		ArrayList<String> a = new ArrayList<String>();

		File execlLoc = new File("C:\\Users\\acer\\Downloads\\demodata.xlsx");
		FileInputStream fis = new FileInputStream(execlLoc);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i); // sheet is collection of rows
				Iterator<Row> rows = sheet.iterator();
				Row row = rows.next(); // points to first Row ==> TestCasees Data1 Data2 Data3//row is collection of
										// cells

				Iterator<Cell> cells = row.cellIterator();
				int k = 0;
				int column = 0;
				while (cells.hasNext()) {
					Cell value = cells.next();
					System.out.println("header found " + value.toString());

					if (value.toString().trim().equalsIgnoreCase("TestCases")) {
						// desired column
						column = k;
						break;

					}
					k++;

				}
				System.out.println(column);
				// once column header is identified thenn scan entire testcase coulnm to
				// identify puchse testcase row
				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).toString().trim().equalsIgnoreCase("Purchase")) {
						// after you grab purchase testcase for = pull all the data of that row and then
						// feed it to the test case
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							a.add(cv.next().toString());

						}

					}

				}

			}
		}
		return a;
	}

	public static void main(String[] args) throws IOException {

	}

}
