package com.example.project.TestXLs;

import java.lang.reflect.Field;
import java.util.Date;

public class DummyRecord extends ExcelReadWriteEnabler
{
	
	private String rec1;
	private Date rec2;
	private int rec3;
	private float rec4;
	private boolean rec5;
	
	public DummyRecord()
	{
		
	}

	protected Field[] getPrivateFields()
	{
		Field[] fieldNames = this.getClass().getDeclaredFields();
		return fieldNames;
	}
	
	public String[] getHeaders()
	{
		String[] header = {"rec1","rec2","rec3","rec4","rec5"};
		return header;
	}

	public String getRec1() {
		return rec1;
	}

	public Date getRec2() {
		return rec2;
	}

	public void setRec2(Date rec2) {
		this.rec2 = rec2;
	}

	public int getRec3() {
		return rec3;
	}

	public void setRec3(int rec3) {
		this.rec3 = rec3;
	}

	public float getRec4() {
		return rec4;
	}

	public void setRec4(float rec4) {
		this.rec4 = rec4;
	}

	public boolean isRec5() {
		return rec5;
	}

	public void setRec5(boolean rec5) {
		this.rec5 = rec5;
	}

	public void setRec1(String rec1) {
		this.rec1 = rec1;
	}

	
}
