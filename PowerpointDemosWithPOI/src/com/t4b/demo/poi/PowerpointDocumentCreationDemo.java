package com.t4b.demo.poi;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class PowerpointDocumentCreationDemo {
	public static void main(String[] args) {
		XMLSlideShow xmlSlideShow = new XMLSlideShow();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.pptx");
			xmlSlideShow.write(outputStream);
			xmlSlideShow.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
