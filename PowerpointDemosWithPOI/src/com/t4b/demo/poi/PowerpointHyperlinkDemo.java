package com.t4b.demo.poi;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PowerpointHyperlinkDemo {

	public static void main(String[] args) {
		XMLSlideShow xmlSlideShow = new XMLSlideShow();
		try {
			OutputStream outputStream = new FileOutputStream("Demo.pptx");
			XSLFSlideMaster xslfSlideMaster = xmlSlideShow.getSlideMasters().get(0);
			XSLFSlideLayout xslfSlideLayout = xslfSlideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
			XSLFSlide xslfSlide = xmlSlideShow.createSlide(xslfSlideLayout);
			XSLFTextShape textShapeTitle = xslfSlide.getPlaceholder(0);
			textShapeTitle.setText("It's a Hyperlink");
			XSLFTextShape textShapeBody = xslfSlide.getPlaceholder(1);
			textShapeBody.clearText();
			XSLFTextRun r = textShapeBody.addNewTextParagraph().addNewTextRun();
			r.setText("It's another Hyperlink");
			XSLFHyperlink link = r.createHyperlink();
			link.setAddress("http://www.abc.com");
			xmlSlideShow.write(outputStream);
			xslfSlide.clear();
			xmlSlideShow.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
