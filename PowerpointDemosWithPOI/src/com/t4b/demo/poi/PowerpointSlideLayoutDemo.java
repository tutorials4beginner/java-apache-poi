package com.t4b.demo.poi;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PowerpointSlideLayoutDemo {
	public static void main(String[] args) {
		XMLSlideShow xmlSlideShow = new XMLSlideShow();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.pptx");
			XSLFSlideMaster xslfSlideMaster = xmlSlideShow.getSlideMasters().get(0);
			XSLFSlideLayout xslfSlideLayout = xslfSlideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
			XSLFSlide xslfSlide = xmlSlideShow.createSlide(xslfSlideLayout);
			XSLFTextShape textShapeTitle = xslfSlide.getPlaceholder(0);
			textShapeTitle.setText("Demo Title");
			XSLFTextShape textShapeBody = xslfSlide.getPlaceholder(1);
			textShapeBody.clearText();
			textShapeBody.addNewTextParagraph().addNewTextRun().setText("This is a Demo Paragraph.");
			outputStream.close();
			xslfSlide.clear();
			xmlSlideShow.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
