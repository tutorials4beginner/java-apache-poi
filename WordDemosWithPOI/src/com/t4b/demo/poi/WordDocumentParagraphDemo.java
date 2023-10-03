package com.t4b.demo.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocumentParagraphDemo {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XWPFDocument doc = new XWPFDocument();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.doc");
			XWPFParagraph xwpfParagraph = doc.createParagraph();
			XWPFRun xwpfRun = xwpfParagraph.createRun();
			xwpfRun.setText("Dummy paragraph!");
			doc.write(outputStream);
			doc.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
