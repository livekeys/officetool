package com.livekeys.officetool.pptutil;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraphProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAlignType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextFontAlignType;
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

    /**
     * 从幻灯片中获取图表
     * @param slide
     * @return
     */
    public XSLFChart getChartFromSlide(XSLFSlide slide) {
        for (POIXMLDocumentPart relation : slide.getRelations()) {
            if (relation instanceof XSLFChart) {
                return (XSLFChart) relation;
            }
        }
        return null;
    }

    /**
     * 从幻灯片中获取图表列表
     * @param slide
     * @return
     */
    public List<XSLFChart> getAllChartFromSlide(XSLFSlide slide) {
        List<XSLFChart> charts = new ArrayList<XSLFChart>();
        for (POIXMLDocumentPart relation : slide.getRelations()) {
            if (relation instanceof XSLFChart) {
                charts.add((XSLFChart) relation);
            }
        }
        return charts;
    }

    /**
     * 从幻灯片中获取表格
     * @param slide
     * @return
     */
    public XSLFTable getTableFromSlide(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                return (XSLFTable) shape;
            }
        }
        return null;
    }

    /**
     * 从幻灯片中获取表格列表
     * @param slide
     * @return
     */
    public List<XSLFTable> getAllTableFromSlide(XSLFSlide slide) {
        List<XSLFTable> tables = new ArrayList<XSLFTable>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTable) {
                tables.add((XSLFTable) shape);
            }
        }
        return tables;
    }

    /**
     * 从幻灯片中获取文本框
     * @param slide
     * @return
     */
    public XSLFTextBox getTextBoxFromSlide(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextBox) {
                return (XSLFTextBox) shape;
            }
        }
        return null;
    }

    /**
     * 从幻灯片中获取文本框列表
     * @param slide
     * @return
     */
    public List<XSLFTextBox> getAllTextBoxFromSlide(XSLFSlide slide) {
        List<XSLFTextBox> textBoxes = new ArrayList<XSLFTextBox>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextBox) {
                textBoxes.add((XSLFTextBox) shape);
            }
        }
        return textBoxes;
    }

    /**
     * 设置幻灯片段落垂直对齐方式
     * @param paragraph
     * @param vertical
     */
    public void setPargraphVerticalAlign(XSLFTextParagraph paragraph, String vertical) {
        if (vertical == null && "".equals(vertical)) {
            vertical = "auto";
        }

        setCTTextParagraphVerticalAlign(paragraph.getXmlObject(), vertical.toLowerCase());
    }

    // 设置段落垂直对齐
    private void setCTTextParagraphVerticalAlign(CTTextParagraph ctTextParagraph, String verticalStr) {
        CTTextParagraphProperties pPr = ctTextParagraph.getPPr() == null ? ctTextParagraph.addNewPPr() : ctTextParagraph.getPPr();

        switch (verticalStr) {
            case "top" : pPr.setFontAlgn(STTextFontAlignType.T);   break;   // 顶部
            case "baseline" : pPr.setFontAlgn(STTextFontAlignType.BASE);   break;  // 基线对齐
            case "bottom" : pPr.setFontAlgn(STTextFontAlignType.B);   break;    // 底部
            case "center" : pPr.setFontAlgn(STTextFontAlignType.CTR);   break;    // 居中
            default: pPr.setFontAlgn(STTextFontAlignType.AUTO);  // 自动
        }
    }

    /**
     * 设置幻灯片段落水平对齐方式
     * @param paragraph
     * @param horizontal
     */
    public void setParagraphHorizontalAlign(XSLFTextParagraph paragraph, String horizontal) {
        if (horizontal == null && "".equals(horizontal)) {
            horizontal = "auto";
        }

        setCTTextParagraphHorizonAlign(paragraph.getXmlObject(), horizontal.toLowerCase());
    }

    // 设置段落水平对齐方式
    private void setCTTextParagraphHorizonAlign(CTTextParagraph ctTextParagraph, String horizontalStr) {
        CTTextParagraphProperties pPr = ctTextParagraph.getPPr() == null ? ctTextParagraph.addNewPPr() : ctTextParagraph.getPPr();
        switch (horizontalStr) {
            case "left" : pPr.setAlgn(STTextAlignType.L);   break;  // 左对齐
            case "right": pPr.setAlgn(STTextAlignType.R);   break;  // 右对齐
            case "center": pPr.setAlgn(STTextAlignType.CTR);    break;  // 居中
            case "disperse" : pPr.setAlgn(STTextAlignType.DIST);    break;  // 分散对齐
            default: pPr.setAlgn(STTextAlignType.JUST); // 两端对齐
        }
    }

    public void test() {

    }

}
