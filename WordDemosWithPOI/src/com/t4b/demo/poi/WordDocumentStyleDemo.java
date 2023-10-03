package com.t4b.demo.poi;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocumentStyleDemo {

	public static void main(String[] args) {
		XWPFDocument xwpfDocument = new XWPFDocument();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.docx");
			XWPFParagraph paragraph = xwpfDocument.createParagraph();
			XWPFRun xwpfRun = paragraph.createRun();
			xwpfRun.setBold(true);
			xwpfRun.setItalic(true);
			xwpfRun.setText("Demo Text!");
			xwpfRun.addBreak();
			xwpfDocument.write(outputStream);
			xwpfDocument.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
