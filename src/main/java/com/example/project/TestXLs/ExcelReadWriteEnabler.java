package com.example.project.TestXLs;

import java.lang.reflect.Field;

public class ExcelReadWriteEnabler {

	protected Field[] getPrivateFields()
	{
		Field[] fieldNames = this.getClass().getDeclaredFields();
		return fieldNames;
	}
	
	
}
