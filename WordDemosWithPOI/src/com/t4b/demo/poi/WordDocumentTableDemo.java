package com.t4b.demo.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordDocumentTableDemo {

	public static void main(String[] args) {
		XWPFDocument xwpfDocument = new XWPFDocument();
		try (FileOutputStream fileOutputStream = new FileOutputStream(new File("Demo.docx"))) {
			XWPFTable xwpfTable = xwpfDocument.createTable();

			XWPFTableRow xwpfTableRow = xwpfTable.getRow(0);
			xwpfTableRow.getCell(0).setText("Sl.");
			xwpfTableRow.addNewTableCell().setText("Name");
			xwpfTableRow.addNewTableCell().setText("Address");

			xwpfTableRow = xwpfTable.createRow();
			xwpfTableRow.getCell(0).setText("1.");
			xwpfTableRow.getCell(1).setText("Jogn");
			xwpfTableRow.getCell(2).setText("john@abc.com");

			xwpfTableRow = xwpfTable.createRow();
			xwpfTableRow.getCell(0).setText("2.");
			xwpfTableRow.getCell(1).setText("Rich");
			xwpfTableRow.getCell(2).setText("rich@xyz.com");

			xwpfDocument.write(fileOutputStream);
			xwpfDocument.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
