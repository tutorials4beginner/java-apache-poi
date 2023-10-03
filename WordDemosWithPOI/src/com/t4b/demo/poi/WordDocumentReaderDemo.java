package com.t4b.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class WordDocumentReaderDemo {

	public static void readDocFile(String fileName) {
		try {
			File file = new File(fileName);
			FileInputStream fileInputStream = new FileInputStream(file.getAbsolutePath());

			HWPFDocument hwpfDocument = new HWPFDocument(fileInputStream);
			WordExtractor wordExtractor = new WordExtractor(hwpfDocument);
			String[] paragraphs = wordExtractor.getParagraphText();

			System.out.println("Number of paragraph " + paragraphs.length);
			for (String para : paragraphs) {
				System.out.println(para.toString());
			}
			wordExtractor.close();
			fileInputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void readDocxFile(String fileName) {

		try {
			File file = new File(fileName);
			FileInputStream fileInputStream = new FileInputStream(file.getAbsolutePath());
			XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
			List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
			System.out.println("Number of paragraph " + paragraphs.size());
			for (XWPFParagraph paragraph : paragraphs) {
				System.out.println(paragraph.getText());
			}
			xwpfDocument.close();
			fileInputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		readDocxFile("Demo.docx");
		readDocFile("Demo.doc");
	}
}
