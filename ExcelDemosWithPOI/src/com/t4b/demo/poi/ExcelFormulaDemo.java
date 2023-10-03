package com.t4b.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFormulaDemo {
	public static void main(String[] args) {
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Calculate Simple Interest");

		Row header = xssfSheet.createRow(0);
		header.createCell(0).setCellValue("Pricipal");
		header.createCell(1).setCellValue("RoI");
		header.createCell(2).setCellValue("Time");
		header.createCell(3).setCellValue("Interest (P r t)");

		Row dataRow = xssfSheet.createRow(1);
		dataRow.createCell(0).setCellValue(14500d);
		dataRow.createCell(1).setCellValue(9.25);
		dataRow.createCell(2).setCellValue(3d);
		dataRow.createCell(3).setCellFormula("A2*B2*C2");

		try {
			FileOutputStream fileOutputStream = new FileOutputStream(new File("formulaDemo.xlsx"));
			xssfWorkbook.write(fileOutputStream);
			fileOutputStream.close();
			readSheetWithFormula();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void readSheetWithFormula() {
		try {
			FileInputStream fileInputStream = new FileInputStream(new File("formulaDemo.xlsx"));
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
			FormulaEvaluator formulaEvaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
			XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_FORMULA:
						break;
					}
				}
				System.out.println();
			}
			fileInputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
