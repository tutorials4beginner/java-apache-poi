package com.t4b.demo.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PowerpointSlideImageDemo {
	public static void main(String[] args) {
		XMLSlideShow xmlSlideShow = new XMLSlideShow();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.pptx");
			XSLFSlide xslfSlide = xmlSlideShow.createSlide();
			byte[] pictData = IOUtils.toByteArray(new FileInputStream("test.png"));
			XSLFPictureData xslfPictureData = xmlSlideShow.addPicture(pictData, XSLFPictureData.PictureType.PNG);
			xslfSlide.createPicture(xslfPictureData);
			outputStream.close();
			xslfSlide.clear();
			xmlSlideShow.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
