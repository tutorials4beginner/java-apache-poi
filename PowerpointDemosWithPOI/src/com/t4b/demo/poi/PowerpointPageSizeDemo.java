package com.t4b.demo.poi;

import java.awt.Dimension;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class PowerpointPageSizeDemo {
	public static void main(String[] args) {

		try {
			XMLSlideShow xmlSlideShow = new XMLSlideShow();
			Dimension dimension = xmlSlideShow.getPageSize();
			int width = dimension.width;
			int height = dimension.height;
			System.out.println("Width: " + width);
			System.out.println("Height: " + height);
			xmlSlideShow.setPageSize(new java.awt.Dimension(1024, 768));
			java.awt.Dimension newpgsize = xmlSlideShow.getPageSize();
			System.out.println("After resize!");
			System.out.println("Width: " + newpgsize.width);
			System.out.println("Height: " + newpgsize.height);
			xmlSlideShow.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
