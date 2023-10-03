package com.t4b.demo.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class PowerpointSlideDeletionDemo {
	public static void main(String[] args) {
		try {
			XMLSlideShow xmlSlideShow = new XMLSlideShow(new FileInputStream("Demo.pptx"));
			xmlSlideShow.removeSlide(0);
			OutputStream outputStream = new FileOutputStream("Demo.pptx");
			xmlSlideShow.write(outputStream);
			xmlSlideShow.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
