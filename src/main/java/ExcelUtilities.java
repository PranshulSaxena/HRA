																																																																												package com.LIMS.GenericLibrary;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class ExcelUtilities {
	/**
	 * This method is used for
	 * @param SheetName
	 * @param RowNo
	 * @param ColumnNo
	 * @return
	 * @throws Throwable
	 * @throws IOException
	 */
	
	public String readDataIntoExcel(String SheetName,int RowNo,int ColumnNo) throws Throwable, IOException {
	
	FileInputStream ft = new FileInputStream(IPathConstants.excelPath);
	Workbook wt = WorkbookFactory.create(ft);
	Sheet st = wt.getSheet(SheetName);
	Row rw = st.getRow(RowNo);
	Cell cell = rw.getCell(ColumnNo);
	String value = cell.getStringCellValue();
	return value;

}
	public void writeDataExcel(String SheetName,int RowNo,int ColumnNo,String data) throws EncryptedDocumentException, IOException
	{
		FileInputStream ft = new FileInputStream(IPathConstants.excelPath);
		Workbook wt = WorkbookFactory.create(ft);
		Sheet st = wt.getSheet(SheetName);
		Row rw = st.getRow(RowNo);
		Cell cell = rw.getCell(ColumnNo);
		cell.setCellValue(data);
		
		FileOutputStream fi = new FileOutputStream(IPathConstants.excelPath);
		wt.write(fi);
		
	}
	/**
	 * This method is used for get the last
	 * @param SheetName
	 * @return
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
	public int getLastRowNum(String SheetName) throws EncryptedDocumentException, IOException
	{
		FileInputStream ft = new FileInputStream(IPathConstants.excelPath);
		Workbook wt = WorkbookFactory.create(ft);
		Sheet st = wt.getSheet(SheetName);
		int row = st.getLastRowNum();
		return row;
	
	}
	/**
	 * This method is used for get the last cell 
	 * @param SheetName
	 * @return
	 * @throws IOException
	 */
	public int getLastCell(String SheetName) throws  IOException
	{
		FileInputStream ft = new FileInputStream(IPathConstants.excelPath);
		Workbook wt = WorkbookFactory.create(ft);
		Sheet st = wt.getSheet(SheetName);
		Row rw = st.getRow(0);
		int cl = rw.getLastCellNum();
		return cl;
	}
//	public void arrayList (String sheetName,WebDriver driver) throws EncryptedDocumentException, Throwable
//	{
//		FileInputStream fu = new FileInputStream(IPathConstants.excelPath);
//		Workbook wt = WorkbookFactory.create(fu);
//		Sheet st = wt.getSheet(sheetName);
//		Row rw = st.getRow(0);
//		 Row rw1 = st.getRow(1);
//		// Cell cel = rw1.getCell(0);
//		int count =rw1.getLastCellNum();
//	    //int count =st.getRow(1).getLastCellNum();
//		for(int i=0;i<count;i++)
//		{
//			/*String key =st.getRow(0).getCell(i).getStringCellValue();
//			String value=st.getRow(1).getCell(i).getStringCellValue();
//			*/
//			String key =rw.getCell(i).getStringCellValue();
//			String value = rw1.getCell(i).getStringCellValue();
//			
//			driver.findElement(By.name(key)).sendKeys(value);
//			
//		}
//	
	public Map<String,String>getList(String sheetName) throws Throwable{
		FileInputStream fu = new FileInputStream(IPathConstants.excelPath);
    	Workbook wt = WorkbookFactory.create(fu);
		Sheet st = wt.getSheet(sheetName);
		int count =st.getRow(1).getLastCellNum();
		
		Map<String,String> map = new HashMap<String,String>();
		for(int i =0;i<count;i++)
		{
			String key = st.getRow(0).getCell(i).getStringCellValue();
			String value=st.getRow(1).getCell(i).getStringCellValue();
			map.put(key, value);
		}
		return map;
		
	}
	
	public String readDataFromHouseIntoExcel(String SheetName,int RowNo,int ColumnNo) throws Throwable, IOException {
		
		FileInputStream ft = new FileInputStream(IPathConstants.xcelPath);
		Workbook wt = WorkbookFactory.create(ft);
		Sheet st = wt.getSheet(SheetName);
		Row rw = st.getRow(RowNo);
		Cell cell = rw.getCell(ColumnNo);
		String value = cell.getStringCellValue();
		return value;

	}
	public Object[][] readMultipleData(String SheetName) throws Throwable
	{
		FileInputStream fe = new FileInputStream(IPathConstants.excelPath);
		Workbook wf = WorkbookFactory.create(fe);
		Sheet sh = wf.getSheet(SheetName);
		int lastRow = sh.getLastRowNum()+1;
		int lastCell = sh.getRow(0).getLastCellNum();
		
		Object[][]ob=new Object[lastRow][lastCell];
		for(int i=0;i<lastRow;i++)
		{
			for(int j=0;j<lastCell;j++)
			{
				ob[i][j]=sh.getRow(i).getCell(j).getStringCellValue();
			}
		}
		return ob;
		
	}
	
	public Object[][] DataCellWise(String sheetName) throws Throwable
	{
		FileInputStream fis = new FileInputStream(IPathConstants.excelPath);
		Workbook wc = WorkbookFactory.create(fis);
		Sheet sh = wc.getSheet(sheetName);
		int lastRow = sh.getLastRowNum();
     	int lastCell = sh.getRow(0).getLastCellNum();
    	
    	Object [][]ob= new Object[lastRow][lastCell];
    	for(int i=0;i<lastCell;i++)
		{
			for(int j=0;j<lastRow;j++)
			{
				ob[i][j]=sh.getRow(i).getCell(j).getStringCellValue();
			}
		}
		return ob;
	
	}
	}
	
	
