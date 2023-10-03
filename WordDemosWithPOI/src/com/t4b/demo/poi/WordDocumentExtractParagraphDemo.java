package com.t4b.demo.poi;

import java.io.FileInputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class WordDocumentExtractParagraphDemo {

	public static void main(String[] args) {
		try {
			FileInputStream fileInputStream = new FileInputStream("Demo.docx");
			XWPFDocument xwpfDocument = new XWPFDocument(OPCPackage.open(fileInputStream));
			java.util.List<XWPFParagraph> xwpfParagraphs = xwpfDocument.getParagraphs();
			for (XWPFParagraph xwpfParagraph : xwpfParagraphs) {
				System.out.println(xwpfParagraph.getText());
			}
			xwpfDocument.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
