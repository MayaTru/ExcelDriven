package org.maya.ExcelDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class DataDrivenTest {
	public ArrayList<String> getData(String testcasename) throws IOException {
		ArrayList<String> alst = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C://JarsForTestAut//TestBook.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetscount = workbook.getNumberOfSheets();
		for (int i=0;i<sheetscount;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				XSSFSheet sheet =  workbook.getSheetAt(i);
				//Identify TestCases coloum by scanning the entire 1st row
				Iterator<Row> rows = sheet.rowIterator();
				Row firstrow = rows.next();
				Iterator<Cell> cell = firstrow.cellIterator();
				int k = 0;
				int coloumn=0;
				while(cell.hasNext()) {
					Cell value = cell.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						  coloumn=k;
					}
					k++;
				}
				while(rows.hasNext()) {
					Row r = rows.next();
					if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcasename)) {
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext()) {
							Cell c = cv.next();
							String t1 = c.getCellType().toString();
							//System.out.println(t1);
							
							if(t1.equalsIgnoreCase("STRING")) { 
								alst.add(c.getStringCellValue());
							}
							else if(t1.equalsIgnoreCase("NUMERIC")) { 
								alst.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
		}
		return alst;
	}
	
	public static void main(String[] args) throws IOException {
		DataDrivenTest dobj = new DataDrivenTest();
		ArrayList<String> lst1 = dobj.getData("Purchase");
		System.out.println(lst1.get(0));
		System.out.println(lst1.get(1));
		System.out.println(lst1.get(2));
		System.out.println(lst1.get(3));
		
	}
}
