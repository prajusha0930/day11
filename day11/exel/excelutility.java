package com.example.exel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class excelutility {
	public static void main(String args[]) throws IOException{
	
		String location="C:\\Users\\PROBOOK\\Desktop\\exel\\praju.xlsx";
		FileInputStream fis=new FileInputStream(location);
		//creating a work book object
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		//locating the sheet
		XSSFSheet sheet=workbook.getSheetAt(0);
		//getting total max rows
		int total=sheet.getPhysicalNumberOfRows();
		System.out.print("Total number of rows: "+total);
		int column=sheet.getRow(0).getLastCellNum();
		System.out.println("Total column: "+column);
		//getting rows and Column/
		for(int i=1;i<total;i++) {
			XSSFRow row =sheet.getRow(i);
			for(int j=0;j<column;j++)
			{
				XSSFCell cell=row.getCell(j);	
				System.out.println(cell.getNumericCellValue());
			}
		 }
		
		
	}	
}
