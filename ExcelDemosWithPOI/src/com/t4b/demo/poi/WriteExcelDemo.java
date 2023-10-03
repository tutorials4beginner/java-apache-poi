package com.t4b.demo.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelDemo {
	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("Book Information");

		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "ISBN", "NAME", "AUTHOR" });
		data.put("2", new Object[] { "01-8978-76366", "Data Structure", "Anand" });
		data.put("3", new Object[] { "01-8933-74566", "C Programming", "Richi" });
		data.put("4", new Object[] { "02-6366-89788", "Java", "John" });

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objects = data.get(key);
			int cellnum = 0;
			for (Object o : objects) {
				Cell cell = row.createCell(cellnum++);
				if (o instanceof String)
					cell.setCellValue((String) o);
				else if (o instanceof Integer)
					cell.setCellValue((Integer) o);
			}
		}
		try {
			FileOutputStream fileOutputStream = new FileOutputStream(new File("Demo.xlsx"));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
