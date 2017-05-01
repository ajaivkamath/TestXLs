package com.example.project.TestXLs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * * Sample Java program to read and write Excel file in Java using Apache POI *
 */

public class XLSReadWriter {

	private short[] headerBgColour = {24,29,30,31,32,36,37,38,41,42,43,44,45,51,52,53,58,60,61,62};
	private short[] headerFontColour ={0,1,0,0,1,1,1,1,0,0,0,0,0,0,0,0,1,1,1,1};
	public int iRecCount = 0;
	
	
	public List<Object> uploadXLSX(File inputFile, Class<?> className) {
		// For storing data into CSV files
		List<Object> dataObjectList = new ArrayList<Object>();
		Object myClassObj = null;
		Method method = null;
		Field[] fields = null;
		
		try
		{
			myClassObj = className.newInstance();
			method = myClassObj.getClass().getDeclaredMethod("getPrivateFields");
//					getMethod("getPrivateFields");
			fields = (Field[]) method.invoke(myClassObj);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		
		int iCount = 0;

		StringBuffer data = new StringBuffer();
		try {

			// Get the workbook object for XLSX file
			@SuppressWarnings("resource")
			XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(inputFile));

			// Get first sheet from the workbook
			XSSFSheet sheet = wBook.getSheetAt(0);
			Row row;
			Cell cell;
			int iRow = 0;

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			row = rowIterator.next();//Skip firstRow
			
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				iRow++;
				iCount++;
				Object instance = className.newInstance();

				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator(); 

				for (Field field : fields) {

					if (cellIterator.hasNext()) {
						cell = cellIterator.next();
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							set(instance, field.getName(), cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							set(instance, field.getName(),convertTypeData(cell.getNumericCellValue(),field) );
							break;
						case Cell.CELL_TYPE_STRING:
							set(instance, field.getName(), convertTypeData(cell.getStringCellValue(),field));
							break;
						case Cell.CELL_TYPE_BLANK:
							set(instance, field.getName(), null);
							break;
						default:
							set(instance, field.getName(), null);

						}
					} 
					else 
					{
						System.out.println("Row# " + iRow + " : EXCEPTION - Cell found empty from field "
								+ field.getName() + ". Reading cells skipped. ");
						break;
					}

				}
				dataObjectList.add(instance);

			}

			return dataObjectList;

		} 
		catch (Exception ioe) 
		{
			ioe.printStackTrace();
		}
		return null;
	}

	public List<Object> uploadXLS(File inputFile, Class<?> className) {
		// For storing data into CSV files
		List<Object> dataObjectList = new ArrayList<Object>();
		Object myClassObj = null;
		Method method = null;
		Field[] fields = null;
		
		try
		{
			myClassObj = className.newInstance();
			method = myClassObj.getClass().getDeclaredMethod("getPrivateFields");
//					getMethod("getPrivateFields");
			fields = (Field[]) method.invoke(myClassObj);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		
		int iCount = 0;

		StringBuffer data = new StringBuffer();
		try {

			// Get the workbook object for XLSX file
			@SuppressWarnings("resource")
			HSSFWorkbook wBook = new HSSFWorkbook(new FileInputStream(inputFile));

			// Get first sheet from the workbook
			HSSFSheet sheet = wBook.getSheetAt(0);
			Row row;
			Cell cell;
			int iRow = 0;

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			row = rowIterator.next();//Skip firstRow
			
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				iRow++;
				iCount++;
				Object instance = className.newInstance();

				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator(); 

				for (Field field : fields) {

					if (cellIterator.hasNext()) {
						cell = cellIterator.next();
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							set(instance, field.getName(), convertTypeData(cell.getBooleanCellValue(),field));
							break;
						case Cell.CELL_TYPE_NUMERIC:
							set(instance, field.getName(), convertTypeData(cell.getNumericCellValue(),field) );
							break;
						case Cell.CELL_TYPE_STRING:
							set(instance, field.getName(), convertTypeData(cell.getStringCellValue(),field) );
							break;
						case Cell.CELL_TYPE_BLANK:
							set(instance, field.getName(), null);
							break;
						default:
							set(instance, field.getName(), null);

						}
					} 
					else 
					{
						System.out.println("Row# " + iRow + " : EXCEPTION - Cell found empty from field "
								+ field.getName() + ". Reading cells skipped. ");
						break;
					}

				}
				dataObjectList.add(instance);

			}

			return dataObjectList;

		} 
		catch (Exception ioe) 
		{
			ioe.printStackTrace();
		}
		return null;
	}

	
	public void writeXLSX(List<Object> dataObjectList, File outputFile) {

		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			// Get the workbook object for XLSX file
			XSSFWorkbook wBook = new XSSFWorkbook();

			// Get first sheet from the workbook
			XSSFSheet sheet = wBook.createSheet();

			// Iterate through each rows from first sheet
			int iRow = sheet.getLastRowNum();
			int cellnum = 0;
			int headerColumnCount = 0;
			
			try 
			{
				Object myClassObj =  (dataObjectList.get(0)).getClass().newInstance();
				Method method = myClassObj.getClass().getMethod("getHeaders");
				String[] fieldHeaders = (String[]) method.invoke(myClassObj);
				XSSFRow row = sheet.createRow(iRow++);
				headerColumnCount = fieldHeaders.length;
				int bgColorIndex = (new Random()).nextInt(headerBgColour.length);
				short xlsHeaderBgColour = headerBgColour[bgColorIndex]; 
				short xlsHeaderFontColour = headerFontColour[bgColorIndex];
				
				for (String fieldHeader: fieldHeaders) 
				{
					Cell cell = row.createCell(cellnum++);
					CellStyle cellStyle = setBordersMethod(wBook.createCellStyle());
					cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND); 
					cellStyle.setFillForegroundColor( xlsHeaderBgColour );
					
					cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
				    XSSFFont font = wBook.createFont();
				    font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				    font.setColor(xlsHeaderFontColour);
				    cellStyle.setFont(font);
					cell.setCellValue(fieldHeader.toUpperCase());
					cell.setCellStyle(cellStyle);
				}
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			
			for (Object rowObj : dataObjectList) {
				XSSFRow row = sheet.createRow(iRow++);
				Field[] fields =  rowObj.getClass().getDeclaredFields();
				cellnum = 0;
				
				cellnum = 0;
				for (Field field : fields) 
				{
					Cell cell = row.createCell(cellnum++);
					CellStyle cellStyle = setBordersMethod(wBook.createCellStyle());

					if (field.getType().toString().equals("class java.lang.String"))
					{
						cell.setCellValue((String) get(rowObj, field.getName()));
					}
					else if (field.getType().toString().equals("class java.lang.Boolean") || field.getType().toString().equals("boolean"))
					{
						
						cell.setCellValue((Boolean) get(rowObj, field.getName()));
					}						
					else if (field.getType().toString().equals("class java.lang.Integer") || field.getType().toString().equals("int"))
					{
						cell.setCellValue((Integer) get(rowObj, field.getName()));
					}
					else if (field.getType().toString().equals("class java.lang.Double") || field.getType().toString().equals("double")  )
					{
						XSSFDataFormat df = wBook.createDataFormat();
						cell.setCellValue((Double) get(rowObj, field.getName()));
						cellStyle.setDataFormat(df.getFormat("#.##")); // custom number format
					}
					else if (field.getType().toString().equals("class java.lang.Float") || field.getType().toString().equals("float"))
					{
						XSSFDataFormat df = wBook.createDataFormat();
						cell.setCellValue(get(rowObj, field.getName()).toString()) ;
						cellStyle.setDataFormat(df.getFormat("#.##")); // custom number format
					}
					else if (field.getType().toString().equals("class java.lang.Long") || field.getType().toString().equals("long"))
					{
						cell.setCellValue(get(rowObj, field.getName()).toString()) ;
						
					}
					else if (field.getType().toString().equals("class java.util.Date"))
					{
						XSSFDataFormat df = wBook.createDataFormat();
						cellStyle.setDataFormat(df.getFormat("dd-MM-yyyy"));
						cell.setCellValue((Date) get(rowObj, field.getName()));
					}
					else
					{
						cell.setCellValue((String)get(rowObj, field.getName()));
					}
					cell.setCellStyle(cellStyle);
				}

			}
			
			for (int i = 0;i < headerColumnCount;i++)
				
			{
				wBook.getSheetAt(0).autoSizeColumn(i);
				if (wBook.getSheetAt(0).getColumnWidth(i) < 2048)
					wBook.getSheetAt(0).setColumnWidth(i, 2048); 
			}

			wBook.write(fos);
			System.out.println("Writing on Excel file Finished ...");
			fos.close();
			wBook.close();
		} catch (FileNotFoundException fe) {
			fe.printStackTrace();
		} catch (IOException ie) {
			ie.printStackTrace();
		}

	}

	public boolean set(Object object, String fieldName, Object fieldValue) {
		Class<?> clazz = object.getClass();
		while (clazz != null) {
			try {
				Field field = clazz.getDeclaredField(fieldName);
				field.setAccessible(true);
				field.set(object, fieldValue);
				return true;
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
				clazz = clazz.getSuperclass();
			} catch (Exception e) {
				e.printStackTrace();
				throw new IllegalStateException(e);
			}
		}
		return false;
	}

	public void writeXLS(List<Object> dataObjectList, File outputFile) {

		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			// Get the workbook object for XLSX file
			HSSFWorkbook wBook = new HSSFWorkbook();

			// Get first sheet from the workbook
			HSSFSheet sheet = wBook.createSheet();

			// Iterate through each rows from first sheet
			int iRow = sheet.getLastRowNum();
			int cellnum = 0;
			int headerColumnCount = 0;
			
			try 
			{
				Object myClassObj =  (dataObjectList.get(0)).getClass().newInstance();
				Method method = myClassObj.getClass().getMethod("getHeaders");
				String[] fieldHeaders = (String[]) method.invoke(myClassObj);
				HSSFRow row = sheet.createRow(iRow++);
				headerColumnCount = fieldHeaders.length;
				int bgColorIndex = (new Random()).nextInt(headerBgColour.length);
				short xlsHeaderBgColour = headerBgColour[bgColorIndex]; 
				short xlsHeaderFontColour = headerFontColour[bgColorIndex];
				
				for (String fieldHeader: fieldHeaders) 
				{
					Cell cell = row.createCell(cellnum++);
					CellStyle cellStyle = setBordersMethod(wBook.createCellStyle());
					cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND); 
					cellStyle.setFillForegroundColor( xlsHeaderBgColour );
					
//					cellStyle = setBordersMethod(cellStyle);
					
					cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
				    HSSFFont font = wBook.createFont();
				    font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				    font.setColor(xlsHeaderFontColour);
				    cellStyle.setFont(font);
					cell.setCellValue(fieldHeader.toUpperCase());
					cell.setCellStyle(cellStyle);
				}
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			
			
			for (Object rowObj : dataObjectList) {
				HSSFRow row = sheet.createRow(iRow++);
				Field[] fields =  rowObj.getClass().getDeclaredFields();
				cellnum = 0;
				
				cellnum = 0;
				for (Field field : fields) 
				{
					Cell cell = row.createCell(cellnum++);
					CellStyle cellStyle = setBordersMethod(wBook.createCellStyle());

					if (field.getType().toString().equals("class java.lang.String"))
					{
						cell.setCellValue((String) get(rowObj, field.getName()));
					}
					else if (field.getType().toString().equals("class java.lang.Boolean") || field.getType().toString().equals("boolean"))
					{
						cell.setCellValue((Boolean) get(rowObj, field.getName()));
					}						
					else if (field.getType().toString().equals("class java.lang.Integer") || field.getType().toString().equals("int"))
					{
						cell.setCellValue((Integer) get(rowObj, field.getName()));
					}
					else if (field.getType().toString().equals("class java.lang.Double") || field.getType().toString().equals("double")  )
					{
						HSSFDataFormat df = wBook.createDataFormat();
						cell.setCellValue((Double) get(rowObj, field.getName()));
						cellStyle.setDataFormat(df.getFormat("#.##")); // custom number format
					}
					else if (field.getType().toString().equals("class java.lang.Float") || field.getType().toString().equals("float"))
					{
						HSSFDataFormat df = wBook.createDataFormat();
						cell.setCellValue(get(rowObj, field.getName()).toString()) ;
						cellStyle.setDataFormat(df.getFormat("#.##")); // custom number format
					}
					else if (field.getType().toString().equals("class java.lang.Long") || field.getType().toString().equals("long"))
					{
						cell.setCellValue(get(rowObj, field.getName()).toString()) ;
						
					}
					else if (field.getType().toString().equals("class java.util.Date"))
					{
						HSSFDataFormat df = wBook.createDataFormat();
						cellStyle.setDataFormat(df.getFormat("dd-MM-yyyy"));
						cell.setCellValue((Date) get(rowObj, field.getName()));
					}
					else
					{
						cell.setCellValue((String)get(rowObj, field.getName()));
					}
//					cellStyle = setBordersMethod(cellStyle);
					cell.setCellStyle(cellStyle);
				}

			}
			
			for (int i = 0;i < headerColumnCount;i++)
				
			{
				wBook.getSheetAt(0).autoSizeColumn(i);
				if (wBook.getSheetAt(0).getColumnWidth(i) < 2048)
					wBook.getSheetAt(0).setColumnWidth(i, 2048); 
			}

			wBook.write(fos);
			System.out.println("Writing on Excel file Finished ...");
			fos.close();
			wBook.close();
		} catch (FileNotFoundException fe) {
			fe.printStackTrace();
		} catch (IOException ie) {
			ie.printStackTrace();
		}

	}

	
	private CellStyle setBordersMethod(CellStyle cellStyle) {
		iRecCount++;
		System.out.println(iRecCount);
		cellStyle.setTopBorderColor((short)23);
		cellStyle.setBottomBorderColor((short)23);
		cellStyle.setLeftBorderColor((short)23);
		cellStyle.setRightBorderColor((short)23);
		cellStyle.setBorderTop((short) 1);
		cellStyle.setBorderBottom((short) 1);
		cellStyle.setBorderLeft((short) 1);
		cellStyle.setBorderRight((short) 1);
		return cellStyle;
	}

	@SuppressWarnings("unchecked")
	public <V> V get(Object object, String fieldName) {
		Class<?> clazz = object.getClass();
		while (clazz != null) {
			try {
				Field field = clazz.getDeclaredField(fieldName);
				field.setAccessible(true);
				return (V) field.get(object);
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
				clazz = clazz.getSuperclass();
			} catch (Exception e) {
				e.printStackTrace();
				throw new IllegalStateException(e);
			}
		}
		return null;
	}

	public Date convertStringToDate(String strDate) {

		String[] dateFormatString = {"EEE MMM d HH:mm:ss zzz yyyy","EEE MMM dd yyyy hh:mm aaa","dd/MM/yyyy","dd-MM-yyyy HH:mm:ss","dd-MMM-yyyy","MM dd, yyyy","E, MMM dd yyyy","E, MMM dd yyyy HH:mm:ss"}; 
       DateFormat df;
       
       Date resultDate;
       for (int i = 0;i < dateFormatString.length;i++)
       {
	       try
	       {
	    	   df = new SimpleDateFormat(dateFormatString[i]);
	    	   resultDate = (Date)df.parse(strDate);;
	    	   return resultDate;
	    	   
	       }
	       catch (Exception e)
	       {
	    	   e.printStackTrace();
	       }
       }
       return null;
 	}

	public Object convertTypeData(Object dataObject, Field field) {
		
		if (dataObject==null || field == null)
		{
			return null;
		}
		else if (field.getType().toString().equals("class java.lang.String"))
		{
			if (dataObject.getClass().equals("class java.util.Date"))
			{
				String strDate = dataObject.toString();
				return convertStringToDate(strDate);
			}
			else
				return dataObject.toString();
		}
		else if (field.getType().toString().equals("class java.lang.Boolean") || field.getType().toString().equals("boolean"))
		{
			return (Boolean) dataObject;
		}						
		else if (field.getType().toString().equals("class java.lang.Integer") || field.getType().toString().equals("int"))
		{
			String str = dataObject.toString();
			
			return Integer.parseInt(str.replaceAll("\\.0*$", ""));
		}
		else if (field.getType().toString().equals("class java.lang.Double") || field.getType().toString().equals("double")  )
		{
			return Double.parseDouble(dataObject.toString());
		}
		else if (field.getType().toString().equals("class java.lang.Float") || field.getType().toString().equals("float"))
		{
			return Float.parseFloat(dataObject.toString());
		}
		else if (field.getType().toString().equals("class java.lang.Long") || field.getType().toString().equals("long"))
		{
			return Long.parseLong(dataObject.toString());
		}
		
		else if (field.getType().toString().equals("class java.util.Date"))
		{
			Date startDate = null;
			try
			{
				if (dataObject.getClass().toString().equals("class java.lang.Long") || dataObject.getClass().toString().equals("long"))
				{
					return (HSSFDateUtil.getJavaDate(Long.parseLong(dataObject.toString())));
				}
				else if (dataObject.getClass().toString().equals("class java.lang.Double") || dataObject.getClass().toString().equals("double"))
				{
					double dblDate = Double.parseDouble(dataObject.toString());
					return (HSSFDateUtil.getJavaDate(dblDate));
				}
				else
				{
					String dateStr = dataObject.toString(); 
					return convertStringToDate(dateStr);
				}
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			return startDate;
			
		}
		else
		{
			return dataObject.toString();
		}
	}

	
}
