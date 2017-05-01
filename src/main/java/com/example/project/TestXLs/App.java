package com.example.project.TestXLs;

import java.io.File;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.util.Date;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
//        System.out.println( "Hello World!" );
        DummyRecord d = new DummyRecord();
		File inputFile = new File("C:\\Users\\Ajai V Kamath\\Documents\\Result.xlsx");
		File outputExtractFile = new File("C:\\Users\\Ajai V Kamath\\Documents\\New_Test6.xlsx");
		File inputXLSFile = new File("C:\\Users\\Ajai V Kamath\\Documents\\Result.xls");
		File outputExtractXLSFile = new File("C:\\Users\\Ajai V Kamath\\Documents\\New_Test6.xls");
		
		XLSReadWriter xlsWriter = new XLSReadWriter();
		List<Object> dataObjectList = xlsWriter.uploadXLSX(inputFile, DummyRecord.class);
		//xlsWriter.writeXLSX(dataObjectList, outputExtractFile);

		//List<Object> dataObjectList2 = xlsWriter.uploadXLS(inputXLSFile, DummyRecord.class);
		xlsWriter.writeXLS(dataObjectList, outputExtractXLSFile);

		
        Field[] fields = d.getPrivateFields();
        System.out.println(fields.length);
        
//    	String str = "String";
//    	Object obj = (String) str;
//    	
//    	Object obj1 = XLSReadWriter.convertTypeData(obj,d.getPrivateFields()[0]);
//    	String str1 = obj1.toString();
//    	
//    	d.setRec1(str1);
//    	
//    	System.out.println(d.getRec1());
//    	
//
//    	int str2 = 1;
//    	Object obj2 = (int) str2;
//    	
//    	Object obj3 = XLSReadWriter.convertTypeData(obj2,d.getPrivateFields()[3]);
//    	int str3 = Integer.parseInt(obj3.toString());
//    	
//    	d.setRec3(str3);
//    	
//    	System.out.println(d.getRec3());
//
//    	Date str4 =  new Date();
//    	Object obj4 =  str4 ;
//    	
//    	System.out.println(DateFormat.getDateInstance().toString());
//    	
//    	System.out.println(XLSReadWriter.convertStringToDate(str4.toString()));
//    	
//    	Object objx =  (long) 1346524199000l;
//    	Object objy =  (double) 1346524199000l;
//    	
//    	Object obj5 = XLSReadWriter.convertTypeData(obj4,d.getPrivateFields()[1]);
//    	Object obj6 = XLSReadWriter.convertTypeData(objx,d.getPrivateFields()[1]);
//    	Object obj7 = XLSReadWriter.convertTypeData(objy,d.getPrivateFields()[1]);
//    	
//    	Date str4_1 = (Date) XLSReadWriter.convertTypeData(obj5,d.getPrivateFields()[1]);
//    	Date str4_x = (Date) XLSReadWriter.convertTypeData(obj6,d.getPrivateFields()[1]);
//    	Date str4_y = (Date) XLSReadWriter.convertTypeData(obj7,d.getPrivateFields()[1]);
//    	
//    	d.setRec2(str4_1);
//    	
//    	System.out.println(d.getRec2());
//    	
//    	d.setRec2(str4_x);
//
//    	System.out.println(d.getRec2());
//    	
//    	d.setRec2(str4_y);
//
//    	System.out.println(d.getRec2());
    	
//    	Date str5 =  
//    			
//    			Integer.parseInt(obj3.toString());
//    	
//    	d.setRec3(str5);
//    	
//    	System.out.println(d.getRec3());
//    	System.out.println(obj1.getClass());
//    	System.out.println(obj2.getClass());
//    	System.out.println(obj3.getClass());
//    	System.out.println(obj4.getClass());
//    	System.out.println(objx.getClass());
//    	System.out.println(objy.getClass());
//    	System.out.println(str4);
    	
    	
    	
    }
}


