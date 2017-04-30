package com.example.project.TestXLs;

import java.lang.reflect.Field;

public abstract class ExcelReadWriteEnabler {

	protected abstract Field[] getPrivateFields();
//	{
//		Field[] fieldNames = this.getClass().getDeclaredFields();
//		return fieldNames;
//	}
	
	protected abstract String[] getHeaders();
//	{
//		return null;
//	}
	
	
}
