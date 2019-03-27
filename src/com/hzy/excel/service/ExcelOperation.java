package com.hzy.excel.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperation {
	// path：excel 存放路径; name:excel全名，例如：abc.xlsx 或 abc.xls
	public ExcelSheet readExcel(String path, String name) {
		if (path == null || name == null)
			return null;
		if (name.endsWith("xls"))
			return readXls(path + name);
		else if (name.endsWith("xlsx"))
			return readXlsx(path + name);
		return null;
	}

	public boolean writeExcel(String path, String name) {
		if (path == null || name == null)
			return true;
		if (name.endsWith("xls"))
			return true;
		else if (name.endsWith("xlsx"))
			return true;
		return false;
	}
/*path:路径和文件全名（有后缀）
 * data：<String, List<List<String>>>:第一个泛型是excel分页（sheet）名字
 * 第二个泛型的外层list代表行， 里面list代表列，装有的string就是列的内容
 */
	public boolean writeXls(String path, Map<String, List<List<String>>> data) {
		if (path != null) {
			Workbook wb = null;
			if (path.endsWith(".xls")) {
				wb = new HSSFWorkbook();
			} else if (path.endsWith(".xlsx")) {
				wb = new XSSFWorkbook();
			} else {
				try {
					throw new Exception("当前文件不是excel文件");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			try(FileOutputStream fileOutputStream = new FileOutputStream(new File(path));)
			{
				for (String sheetName : data.keySet()) {
					Sheet sheet = wb.createSheet(sheetName);
					List<List<String>> rowList = data.get(sheetName);
					for (int i = 0; i < rowList.size(); i++) {
						List<String> cellList = rowList.get(i);
						Row row = sheet.createRow(i);
						for (int j = 0; j < cellList.size(); j++) {
							Cell cell = row.createCell(j);
							cell.setCellValue(cellList.get(j));
						}
					}
				}
				wb.write(fileOutputStream);
			} catch (Exception e) {
				e.printStackTrace();
			} 
		//		finally {
//				if (wb != null) {
//					// wb.close();
//				}
//			}
		}
		return true;
	}

	// path：excel 存放路径，仅接受xlsx格式的excel
	private ExcelSheet readXlsx(String path) {

		XSSFWorkbook xssfWorkbook = null;
		try {
			xssfWorkbook = new XSSFWorkbook(new FileInputStream(path));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets();
		// numSheet++)
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		if (xssfSheet == null)
			return null;
		int columns = this.getMaxColumn(xssfSheet); // 获取最大列数
		int rows = xssfSheet.getLastRowNum(); // 获取总行数
		String[] title = this.getTitle(columns);
		String[][] contains = new String[rows + 1][columns];
		// Read the Row
		for (int rowNum = 0; rowNum <= rows; rowNum++) {
			XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			if (xssfRow != null) {
				for (int col = 0; col < columns; col++) {
					XSSFCell xs = xssfRow.getCell(col);
					if (xs != null)
						contains[rowNum][col] = xs.toString();
					else
						contains[rowNum][col] = null;
				}
			}
		}
		String[] element = path.split("\\\\");
		String sheetName = element[element.length - 1];
		return new ExcelSheet(sheetName, title, contains);
	}

	private int getMaxColumn(XSSFSheet xssfSheet) {
		int rows = xssfSheet.getLastRowNum();
		HashSet<Integer> hs = new HashSet<Integer>();
		for (int i = 0; i < rows; i++)
			hs.add(xssfSheet.getRow(i).getPhysicalNumberOfCells());
		return Collections.max(hs).intValue();
	}

	// path：excel 存放路径，仅接受xls格式的excel
	private ExcelSheet readXls(String path) // throws IOException
	{
		HSSFWorkbook hssfWorkbook = null;
		try {
			hssfWorkbook = new HSSFWorkbook(new FileInputStream(path));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// hssfWorkbook.getNumberOfSheets();//获取excel数量
		HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
		if (hssfSheet == null)
			return null;

		// 提取excel表的名字
		String[] element = path.split("\\\\");
		String sheetName = element[element.length - 1];

		int tableColumns = this.getMaxColumn(hssfSheet); // excel 有效总列数,从零开始计数
		int tableRows = hssfSheet.getLastRowNum(); // excel 有效总行数,从零开始计数

		String[] title = this.getTitle(this.getMaxColumn(hssfSheet)); // excel
																		// 字段名
		String[][] contains = new String[tableRows + 1][tableColumns]; // 每行，每列内容

		// Read the Row
		for (int rowNum = 0; rowNum <= tableRows; rowNum++) {
			HSSFRow hssfRow = hssfSheet.getRow(rowNum);
			if (hssfRow != null) {
				for (int i = 0; i < tableColumns; i++) {
					HSSFCell hf = hssfRow.getCell(i);
					if (hf != null)
						contains[rowNum][i] = hf.toString();
					else
						contains[rowNum][i] = null;
				}

			}
		}
		return new ExcelSheet(sheetName, title, contains);
	}

	// 获取最大列数;hss = hssfWorkbook.getSheetAt(numSheet);
	public int getMaxColumn(HSSFSheet hss) {
		int rows = hss.getLastRowNum();
		HashSet<Integer> al = new HashSet<Integer>();
		for (int i = 0; i < rows; i++)
			al.add(hss.getRow(i).getPhysicalNumberOfCells());

		return Collections.max(al).intValue();
	}

	// 获取excel字段; cols:列数
	public String[] getTitle(int cols) {
		String[] str = new String[cols];
		for (int i = 0; i < cols; i++)
			str[i] = CellReference.convertNumToColString(i);
		return str;
	}

}
