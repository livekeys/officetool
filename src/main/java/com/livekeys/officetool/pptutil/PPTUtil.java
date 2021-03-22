package com.livekeys.officetool.pptutil;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PPTUtil {

    private static final Logger logger = LoggerFactory.getLogger(PPTUtil.class);

    private XMLSlideShow pptx;

    public PPTUtil(String filePath) {
        this.readPPT(filePath);
    }

    public XMLSlideShow getPPTX() {
        return pptx;
    }

    // 读取 ppt
    private XMLSlideShow readPPT(String filePath) {
        try {
            this.pptx = new XMLSlideShow(new FileInputStream(new File(filePath)));
            return this.pptx;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    // 写入 ppt
    public void writePPT(String exportPath) {
        try {
            File file = new File(exportPath);
            if (file.exists()) {
                file.delete();
            }
            this.pptx.write(new FileOutputStream(new File(exportPath)));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public XSLFChart getChartFromSlide(XSLFSlide slide) {
        for (POIXMLDocumentPart relation : slide.getRelations()) {
            if (relation instanceof XSLFChart) {
                return (XSLFChart) relation;
            }
        }
        return null;
    }

    public List<XSLFChart> getAllChartFromSlide(XSLFSlide slide) {
        List<XSLFChart> charts = new ArrayList<XSLFChart>();
        for (POIXMLDocumentPart relation : slide.getRelations()) {
            if (relation instanceof XSLFChart) {
                charts.add((XSLFChart) relation);
            }
        }
        return charts;
    }

    public XSLFTable getTableFromSlide(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                return (XSLFTable) shape;
            }
        }
        return null;
    }

    public List<XSLFTable> getAllTableFromSlide(XSLFSlide slide) {
        List<XSLFTable> tables = new ArrayList<XSLFTable>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                tables.add((XSLFTable) shape);
            }
        }
        return tables;
    }

    public XSLFTextBox getTextBoxFromSlide(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextBox) {
                return (XSLFTextBox) shape;
            }
        }
        return null;
    }

    public List<XSLFTextBox> getAllTextBoxFromSlide(XSLFSlide slide) {
        List<XSLFTextBox> textBoxes = new ArrayList<XSLFTextBox>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextBox) {
                textBoxes.add((XSLFTextBox) shape);
            }
        }
        return textBoxes;
    }

    public void test() {
        List<XSLFSlide> slides = pptx.getSlides();

        for (int i = 2; i < slides.size(); i++) {
            XSLFSlide slide = slides.get(i);
            XSLFChart chart = this.getChartFromSlide(slide);
            logger.info(chart.getTitleShape().getText());
	    
        }

    }

}
