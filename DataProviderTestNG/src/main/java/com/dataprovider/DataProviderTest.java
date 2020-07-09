package com.dataprovider;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


public class DataProviderTest
{
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;
	XSSFCell cell1;
	XSSFCell user;
	XSSFCell pass;

	@Test(dataProvider="playerDetailsDP")
	public void PrintPlayerDetails(String username, String password, String firstname, String lastname) {
		System.out.println("Username:"+username);
		System.out.println("Password:"+password);
		System.out.println("Firstname:"+firstname);
		System.out.println("Lastname:"+lastname);
		System.out.println("******************");
	}
	
	@DataProvider(name="playerDetailsDP",parallel=true)
	public Object[][] playerData() throws IOException {
		Object[][] arrayObject = getExcelData(System.getProperty("user.dir")+"\\testdata.xlsx","playerDetails");
		return arrayObject;
	}

	public String[][]  getExcelData(String filePath, String sheetName) throws IOException
	{
		String [] [] arrayExcelData = null;
		File file = null;
		FileInputStream fis = null;
		XSSFWorkbook wb =null;
		try  
		{  
			file = new File(filePath);
			//Load the file xlsx file
			fis = new FileInputStream(file);
			//creating Workbook instance that refers to .xlsx file  
			wb = new XSSFWorkbook(fis);  
			//creating a Sheet object by sheetname
			XSSFSheet sheet = wb.getSheet(sheetName);
			
			//Get the Number of Rows and columns which has data
			int totalNoOfRows = sheet.getPhysicalNumberOfRows();
			int totalNoOfCols = sheet.getRow(0).getPhysicalNumberOfCells();
			
			System.out.println("totalNoOfRows"+totalNoOfRows);
			System.out.println("totalNoOfCols"+totalNoOfCols);
			//Initialize array to read data from excel and return at the end
			arrayExcelData = new String[totalNoOfRows-1][totalNoOfCols];
			//We are taking totalNoOfRows-1 as the limit, since we want to skip the first header row (username,password etc from the sheet)
			for (int row= 0 ; row < totalNoOfRows-1; row++) {
				for (int col= 0; col < totalNoOfCols; col++) {
					XSSFCell cell = sheet.getRow(row).getCell(col);
					//to format data to string content
					DataFormatter df = new DataFormatter();
					arrayExcelData[row][col] = df.formatCellValue(cell);
				}

			}
		}  
		catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		finally {
			if(wb!=null)
				wb.close();
			if(fis!=null)
				fis.close();
		}
		return arrayExcelData;
	} 

}