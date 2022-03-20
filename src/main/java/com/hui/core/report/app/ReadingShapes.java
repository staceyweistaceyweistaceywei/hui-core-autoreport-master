package com.hui.core.report.app;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ReadingShapes {

    public static void main(String args[]) throws IOException {
        //creating a slideshow
        File file = new File("shapes.pptx");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));
        //get slides
        List<XSLFSlide> slide = ppt.getSlides();
        //getting the shapes in the presentation
        System.out.println("Shapes in the presentation:");
        for (int i = 0; i < slide.size(); i++){
            List<XSLFShape> sh = slide.get(i).getShapes();

            XSLFAutoShape rectangle = slide.get(i).createAutoShape();
            rectangle.setShapeType(ShapeType.RECT);
            rectangle.setAnchor(new Rectangle2D.Double(100, 100, 100, 50));
            rectangle.setLineColor(Color.blue);
            rectangle.setFillColor(Color.lightGray);
            for (int j = 0; j < sh.size(); j++){
                //name of the shape
                System.out.println(sh.get(j).getShapeName());
            }
        }
        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);
        out.close();
    }
}
