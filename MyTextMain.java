package com.Anemoi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class MyTextMain {

	public SimpleDateFormat standardFormat = new SimpleDateFormat("dd-MM-yyyy");
//	private static final Logger LOGGER = LoggerFactory.getLogger(com.Anemoi.TextMain.class);
	public static void main(String[] args) throws FileNotFoundException, IOException {

		MyTextMain tm = new MyTextMain();
		tm.test();

	}

	
	
	public Date formatRawDate(Cell dateCell)
	{
		Date dateField = null;
		if(dateCell == null)
			{
			System.out.print("\n Date Value is NULL ::: " );
			}
		else if (dateCell.getCellType() == CellType.STRING) {
			dateField = formatDate(dateCell.getStringCellValue());
			} 
		else if (dateCell.getCellType() == CellType.NUMERIC) {
			dateField = dateCell.getDateCellValue();
			}

		if (dateField != null)
			System.out.print("\n Date Value read successfully ::: " 
						+ standardFormat.format(dateField));	
	
		return dateField ;
	
	}
	
	public Date formatDate(String value) {
		
		String[] dateFormats = {"dd/MM/yyyy","dd.MM.yyyy","dd-MM-yyyy","MM-dd-yyyy" };
		Date myDate;
		for (String format : dateFormats) {
			 SimpleDateFormat sdf = new SimpleDateFormat(format);
			try {
				myDate = sdf.parse(value);
				return myDate;
			} catch (ParseException e) {
			//	e.printStackTrace();
				System.out.print("\nDate parsing failed ::" + value + " is not right for Format " + format);
			}
		}

		System.out.print("\nNo formats matched ::: :" + value);
		return null;
	}

	public void test() throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		FileInputStream file = new FileInputStream("C:\\Users\\Dell\\Downloads\\temp.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0); // Assuming you are reading the first sheet

		DataFormatter formatter = new DataFormatter();

		int rownum = 1;
		System.out.println("sheet" + sheet.getSheetName() + "lastrow" + sheet.getLastRowNum());
		for (Row row1 : sheet) {
			System.out.print("\n Reading Raw  ::: " + (rownum++));
			Cell dateCell = row1.getCell(0);
			
			Date sdate = formatRawDate(dateCell);
			System.out.println("sdate***"+sdate.getTime());
			

			}
		}
}
	


