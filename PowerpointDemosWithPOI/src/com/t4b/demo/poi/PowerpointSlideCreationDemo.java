package com.t4b.demo.poi;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PowerpointSlideCreationDemo {
	public static void main(String[] args) {
		XMLSlideShow xmlSlideShow = new XMLSlideShow();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.pptx");
			XSLFSlide xslfSlide = xmlSlideShow.createSlide();
			xmlSlideShow.write(outputStream);
			xslfSlide.clear();
			xmlSlideShow.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
