package excel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadExcel {

	@Test
	public void readExcel() throws InvalidFormatException, IOException{
		File src=new File("./TestData/Emp.xlsx");
		XSSFWorkbook wbook= new XSSFWorkbook(src);
		XSSFSheet sheet= wbook.getSheetAt(0);
		int rowcount=sheet.getLastRowNum();
		int colcount=sheet.getRow(0).getLastCellNum();
		System.out.println(rowcount);
		System.out.println(colcount);


		for (int i = 1; i <=rowcount; i++) {
			XSSFRow row = sheet.getRow(i);

			for (int j = 0; j < colcount; j++) {
				XSSFCell cell = row.getCell(j);
				System.out.println(cell);
			}
		
	
	/*
		File src=new File("./TestData/Emp.xlsx");
		XSSFWorkbook wbook=new XSSFWorkbook(src);
		XSSFSheet sheet=wbook.getSheetAt(0);
		 sheet.getRow(0);*/
	
	
	
	
	
	 
	}
}
}
