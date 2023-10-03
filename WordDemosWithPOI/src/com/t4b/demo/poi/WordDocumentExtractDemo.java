package com.t4b.demo.poi;

import java.io.FileInputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordDocumentExtractDemo {

	public static void main(String[] args) {
		try {
			FileInputStream fileInputStream = new FileInputStream("Demo.docx");
			XWPFDocument xwpfDocument = new XWPFDocument(OPCPackage.open(fileInputStream));
			XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(xwpfDocument);
			System.out.println(xwpfWordExtractor.getText());
			xwpfWordExtractor.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
