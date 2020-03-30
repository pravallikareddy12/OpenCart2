package com.cts.utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelUtils {

	public static String[][] getSheetIntoStringArray(String fileDetails,String sheetName) throws IOException
	{
		 FileInputStream file=null;
		 XSSFWorkbook book=null;
		 String[][]main=null;
	   try
	   {
		  file = new FileInputStream(fileDetails);
			book = new XSSFWorkbook(file); // it will get workbook from the file
	      //XSSFSheet sheet=book.getSheet("InvalidCredentialsTest");// sheet from the workbook
			XSSFSheet sheet = book.getSheet(sheetName);
			int rowCount = sheet.getPhysicalNumberOfRows();
			int cellCount = sheet.getRow(0).getPhysicalNumberOfCells();
			//System.out.println(rowCount);
			//System.out.println(cellCount);
	
			
			main = new String[rowCount - 1][cellCount];// to exclude "header"-rowCount-1
				for (int i = 1; i < rowCount; i++) // its for excel
			{

				XSSFRow row = sheet.getRow(i);				
				for (int j = 0; j < cellCount; j++) 
				{
					XSSFCell cell = row.getCell(j);
					DataFormatter format = new DataFormatter();
					String cellValue = format.formatCellValue(cell);
					System.out.println(cellValue + " ");
					main[i-1][j] = cellValue;// if excel [1,0] then 'array=[0,0] which is data'

				}
			}
	   }
	   catch(Exception e) {
		   e.printStackTrace();
	   }
	   finally {//close excel
		  book.close();
		  file.close();
	   }
	   return main;
	}

}

