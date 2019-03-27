package com.hzy.excel.service;

//每张Excel表的对象
public class ExcelSheet
{
	private String[] title = null;
	private String[][] contains = null;
	private String name = null;
	
	public ExcelSheet(String[] title, String[][] contains)
	{
		this.title = title;
		this.contains = contains;
	}
	
	public ExcelSheet(String name, String[] title, String[][] contains)
	{
		this(title, contains);
		this.name = name;
	}
	
	public String getName()
	{
		return this.name;
	}
	
	public String[] getTitle()
	{
		return this.title;
	}
	
	public String[][] getContains()
	{
		return this.contains;
	}
	
}
