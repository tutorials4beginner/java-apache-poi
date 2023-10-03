package com.t4b.demo.poi;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordDocumentCreationDemo {

	public static void main(String[] args) {
		XWPFDocument xwpfDocument = new XWPFDocument();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.docx");
			xwpfDocument.write(outputStream);
			xwpfDocument.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
