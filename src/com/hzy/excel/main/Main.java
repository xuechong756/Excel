package com.hzy.excel.main;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import com.hzy.excel.service.ExcelOperation;
import com.hzy.excel.service.ExcelSheet;

public class Main {

	public static void main(String[] args) {

		//write();
		//reade();
		System.out.println("finished");
	}
	
	public static void reade()
	{
		ExcelSheet es = new ExcelOperation().readExcel("E:\\", "fu.xlsx");
		String[][] contains = es.getContains();
		System.out.println(es.getName());
		for(int row = 0; row < contains.length; row++)
		{
			for(int col = 0; col < contains[0].length; col++)
				System.out.print(contains[row][col]);
			System.out.println();
		}
	}
	public static void write()
	{
		Map<String, List<List<String>>> map = new TreeMap<>();
		List<String> col = new ArrayList<>();//这是列
		for(int i = 0; i < 'Z'; i++)
		{
			//放入A - Z 字符
			col.add(String.valueOf(Character.toChars((Integer.valueOf('A' + i).intValue()))));
		}
		List<List<String>> row = new ArrayList<>();//这是行
		row.add(col);
		map.put("sheet1", row);
		ExcelOperation excel = new ExcelOperation();
		excel.writeXls("E:\\fu.xlsx", map);
	}

}
