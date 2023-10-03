package com.t4b.demo.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocumentAlignDemo {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XWPFDocument xwpfDocument = new XWPFDocument();
		try (OutputStream outputStream = new FileOutputStream("Demo.docx")) {
			XWPFParagraph paragraph = xwpfDocument.createParagraph();
			paragraph.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun run = paragraph.createRun();
			run.setText("This is dummy text!");
			xwpfDocument.write(outputStream);
			xwpfDocument.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
