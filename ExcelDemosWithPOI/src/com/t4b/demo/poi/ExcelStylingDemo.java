package com.t4b.demo.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.record.CFRuleRecord.ComparisonOperator;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStylingDemo {

	public static void main(String[] args) throws Exception {
		Workbook workbook = new XSSFWorkbook();

		styleBasedOnValue(workbook.createSheet("Value Based formatting"));
		formatDuplicates(workbook.createSheet("Duplicates formatting"));
		shadeAlt(workbook.createSheet("Alternate rows"));
		expiryInNext30Days(workbook.createSheet("Soon Expired Payments"));

		FileOutputStream fileOutputStream = new FileOutputStream("styleDemo.xlsx");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}

	static void styleBasedOnValue(Sheet sheet) {
		sheet.createRow(0).createCell(0).setCellValue(84);
		sheet.createRow(1).createCell(0).setCellValue(74);
		sheet.createRow(2).createCell(0).setCellValue(50);
		sheet.createRow(3).createCell(0).setCellValue(51);
		sheet.createRow(4).createCell(0).setCellValue(49);
		sheet.createRow(5).createCell(0).setCellValue(41);

		SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();

		ConditionalFormattingRule conditionalFormattingRule1 = sheetConditionalFormatting
				.createConditionalFormattingRule(ComparisonOperator.GT, "70");
		PatternFormatting patternFormatting1 = conditionalFormattingRule1.createPatternFormatting();
		patternFormatting1.setFillBackgroundColor(IndexedColors.BLUE.index);
		patternFormatting1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		ConditionalFormattingRule conditionalFormattingRule2 = sheetConditionalFormatting
				.createConditionalFormattingRule(ComparisonOperator.LT, "50");
		PatternFormatting patternFormatting2 = conditionalFormattingRule2.createPatternFormatting();
		patternFormatting2.setFillBackgroundColor(IndexedColors.GREEN.index);
		patternFormatting2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] cellRangeAddresses = { CellRangeAddress.valueOf("A1:A6") };

		sheetConditionalFormatting.addConditionalFormatting(cellRangeAddresses, conditionalFormattingRule1,
				conditionalFormattingRule2);
	}

	static void formatDuplicates(Sheet sheet) {
		sheet.createRow(0).createCell(0).setCellValue("Code");
		sheet.createRow(1).createCell(0).setCellValue(4);
		sheet.createRow(2).createCell(0).setCellValue(3);
		sheet.createRow(3).createCell(0).setCellValue(6);
		sheet.createRow(4).createCell(0).setCellValue(3);
		sheet.createRow(5).createCell(0).setCellValue(5);
		sheet.createRow(6).createCell(0).setCellValue(8);
		sheet.createRow(7).createCell(0).setCellValue(0);
		sheet.createRow(8).createCell(0).setCellValue(2);
		sheet.createRow(9).createCell(0).setCellValue(8);
		sheet.createRow(10).createCell(0).setCellValue(6);

		SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();

		ConditionalFormattingRule conditionalFormattingRule = sheetConditionalFormatting
				.createConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
		FontFormatting fontFormatting = conditionalFormattingRule.createFontFormatting();
		fontFormatting.setFontStyle(false, true);
		fontFormatting.setFontColorIndex(IndexedColors.BLUE.index);

		CellRangeAddress[] cellRangeAddresses = { CellRangeAddress.valueOf("A2:A11") };

		sheetConditionalFormatting.addConditionalFormatting(cellRangeAddresses, conditionalFormattingRule);
	}

	static void shadeAlt(Sheet sheet) {
		SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();

		ConditionalFormattingRule conditionalFormattingRule = sheetConditionalFormatting
				.createConditionalFormattingRule("MOD(ROW(),2)");
		PatternFormatting patternFormatting = conditionalFormattingRule.createPatternFormatting();
		patternFormatting.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
		patternFormatting.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] cellRangeAddresses = { CellRangeAddress.valueOf("A1:Z100") };

		sheetConditionalFormatting.addConditionalFormatting(cellRangeAddresses, conditionalFormattingRule);
	}

	static void expiryInNext30Days(Sheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("d-mmm"));

		sheet.createRow(0).createCell(0).setCellValue("Date");
		sheet.createRow(1).createCell(0).setCellFormula("TODAY()+29");
		sheet.createRow(2).createCell(0).setCellFormula("A2+1");
		sheet.createRow(3).createCell(0).setCellFormula("A3+1");

		for (int rownum = 1; rownum <= 3; rownum++)
			sheet.getRow(rownum).getCell(0).setCellStyle(cellStyle);

		SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();

		ConditionalFormattingRule conditionalFormattingRule = sheetConditionalFormatting
				.createConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
		FontFormatting fontFormatting = conditionalFormattingRule.createFontFormatting();
		fontFormatting.setFontStyle(false, true);
		fontFormatting.setFontColorIndex(IndexedColors.BLUE.index);

		CellRangeAddress[] cellRangeAddresses = { CellRangeAddress.valueOf("A2:A4") };

		sheetConditionalFormatting.addConditionalFormatting(cellRangeAddresses, conditionalFormattingRule);

		sheet.getRow(0).createCell(1).setCellValue("Dates within the next 30 days are highlighted");
	}
}
