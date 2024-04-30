package com.inetbanking_utility;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataProvider {
	FileInputStream fins;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	
	public ExcelDataProvider(String sheetname) {
		try {
			File fs=new File("./testdata/"+ sheetname+".xlsx");
			fins=new FileInputStream(fs);
			workbook=new XSSFWorkbook(fins);
			sheet=workbook.getSheet(sheetname);
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public int rowCount(String sheetname) {
		return workbook.getSheet(sheetname).getLastRowNum();
	}
	
	public int colCount(String sheetname,int row) {
		return workbook.getSheet(sheetname).getRow(row).getLastCellNum();	
	}
	
	public String fetchStringcellvalue(String sheetname,int row,int col) {
		return workbook.getSheet(sheetname).getRow(row).getCell(col).getStringCellValue();
	}
		
		public int fetchNumericCellvalue(int index,int row,int col) {
			return (int)workbook.getSheetAt(index).getRow(row).getCell(col).getNumericCellValue();
		
		
	}

	public String[][] getExcelTestDate(String sheetname)
	{
		int rows=rowCount(sheetname);
		int col=colCount(sheetname,0);
		
		String[][] data=new String[rows][col];
		for(int i=0;i<rows;i++)
		{
			for(int j=0;j<col;j++) 
			{
				data[i][j]=workbook.getSheet(sheetname).getRow(i+1).getCell(j).toString();
			}
		}
		return data;
		
	}

}
