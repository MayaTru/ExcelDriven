package org.maya.ExcelDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProviderTest {
	
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider = "drivertest")
	public void Test1(String t, String h, String p, String v) {
		System.out.println(t);
		System.out.println(h);
		System.out.println(p);
		System.out.println(v);
	}
	
	//@Test
	public void Test2() throws IOException {
		DataProviderTest o1 = new DataProviderTest();
		Object d[][] = o1.getData();
		System.out.println(d[0][0]);
	}
	
	@DataProvider(name = "drivertest")
	public Object[][] getData() throws IOException {
		FileInputStream fis = new FileInputStream("C://JarsForTestAut//TestBook.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		Object data[][] = null;
		//HashMap<String, String> hmp = new HashMap<String, String>();
		//List<HashMap<String, String>> lhmp;
		int sheetcount = workbook.getNumberOfSheets();
		for (int i=0;i<sheetcount;i++) {
			if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase("Sheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				int rowcount = sheet.getPhysicalNumberOfRows();
				System.out.println(rowcount);
				XSSFRow firstrow = sheet.getRow(0);
				int colcount = firstrow.getLastCellNum();
				System.out.println(colcount);
				data = new Object[rowcount-1][colcount];
				for (int j=0;j<rowcount-1;j++) {
					XSSFRow row = sheet.getRow(j+1);
					for (int k=0;k<colcount;k++) {
						XSSFCell cell = row.getCell(k);
						data[j][k]= formatter.formatCellValue(cell);
					}
				}
			}
		}
		return data;
	}

}
