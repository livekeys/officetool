package com.livekeys.officetool.pptutil;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
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
    public void setParagraphVerticalAlign(XSLFTextParagraph paragraph, String vertical) {
        vertical = this.nullToDefault(vertical, "auto");

        setCTTextParagraphVerticalAlign(paragraph.getXmlObject(), vertical.toLowerCase());
    }

    // 设置段落垂直对齐
    private void setCTTextParagraphVerticalAlign(CTTextParagraph ctTextParagraph, String verticalStr) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
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
        horizontal = this.nullToDefault(horizontal, "auto");

        setCTTextParagraphHorizonAlign(paragraph.getXmlObject(), horizontal.toLowerCase());
    }

    // 设置段落水平对齐方式
    private void setCTTextParagraphHorizonAlign(CTTextParagraph ctTextParagraph, String horizontalStr) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);

        switch (horizontalStr) {
            case "left" : pPr.setAlgn(STTextAlignType.L);   break;  // 左对齐
            case "right": pPr.setAlgn(STTextAlignType.R);   break;  // 右对齐
            case "center": pPr.setAlgn(STTextAlignType.CTR);    break;  // 居中
            case "disperse" : pPr.setAlgn(STTextAlignType.DIST);    break;  // 分散对齐
            default: pPr.setAlgn(STTextAlignType.JUST); // 两端对齐
        }
    }

    // 获取 pPr
    private CTTextParagraphProperties getPPR(CTTextParagraph ctTextParagraph) {
        return ctTextParagraph.getPPr() == null ? ctTextParagraph.addNewPPr() : ctTextParagraph.getPPr();
    }

    /**
     * 设置项目符号的编号
     * @param ctTextParagraph
     * @param lvl
     */
    public void setBulletNum(CTTextParagraph ctTextParagraph, int lvl) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        pPr.setLvl(lvl);
    }

    /**
     * 设置项目符号的颜色
     * @param ctTextParagraph
     * @param colorHex
     */
    public void setBulletColor(CTTextParagraph ctTextParagraph, String colorHex) {
        colorHex = this.nullToDefault(colorHex, "000000");

        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        CTColor buClr = pPr.getBuClr();
        if (pPr.isSetBuClr()) {
            pPr.unsetBuClr();
        }

        CTColor ctColor = pPr.addNewBuClr();
        CTSRgbColor ctsRgbColor = ctColor.addNewSrgbClr();
        ctsRgbColor.setVal(hexToByteArray(colorHex.substring(1)));
    }

    /**
     * 设置段落的行距，单位 磅
     * @param ctTextParagraph
     * @param pounts    磅
     */
    public void setLineSpacingPounts(CTTextParagraph ctTextParagraph, String pounts) {
        pounts = this.nullToDefault(pounts, "1");
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        CTTextSpacing lnSpc = pPr.getLnSpc() == null ? pPr.addNewLnSpc() : pPr.getLnSpc();
        if (lnSpc.isSetSpcPct()) {
            lnSpc.unsetSpcPct();
        }

        CTTextSpacingPoint spcPts = lnSpc.getSpcPts() == null ? lnSpc.addNewSpcPts() : lnSpc.getSpcPts();
        int pts = (int) (Double.valueOf(pounts) * 100);
        spcPts.setVal(pts);
    }

    /**
     * 设置段落的行距，单位倍数
     * @param ctTextParagraph
     * @param multiple 倍数，几倍行距
     */
    public void setLineSpacing(CTTextParagraph ctTextParagraph, Double multiple) {

        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        CTTextSpacing lnSpc = pPr.getLnSpc() == null ? pPr.addNewLnSpc() : pPr.getLnSpc();
        if (lnSpc.isSetSpcPct()) {
            lnSpc.unsetSpcPct();
        }

        CTTextSpacingPercent spcPct = lnSpc.getSpcPct() == null ? lnSpc.addNewSpcPct() : lnSpc.getSpcPct();

        spcPct.setVal(Double.valueOf(multiple * 100000).intValue());
    }

    /**
     * 设置段前间距，单位磅
     * @param ctTextParagraph
     * @param pounts
     */
    public void setCTTextParagraphSpacingBefore(CTTextParagraph ctTextParagraph, String pounts) {
        pounts = this.nullToDefault(pounts, "1");
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        CTTextSpacing ctTextSpacing = pPr.isSetSpcBef() ? pPr.getSpcBef() : pPr.addNewSpcBef();

        if (ctTextSpacing.isSetSpcPct()) {
            ctTextSpacing.unsetSpcPct();
        }

        CTTextSpacingPoint spacing = ctTextSpacing.isSetSpcPts() ? ctTextSpacing.getSpcPts() : ctTextSpacing.addNewSpcPts();
        int pts = (int) (Double.valueOf(pounts) * 100);
        spacing.setVal(pts);
    }

    /**
     * 设置段后间距，单位磅
     * @param ctTextParagraph
     * @param pounts
     */
    public void setCTTextParagraphSpacingAfter(CTTextParagraph ctTextParagraph, String pounts) {
        pounts = this.nullToDefault(pounts, "1");
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        CTTextSpacing ctTextSpacing = pPr.isSetSpcAft() ? pPr.getSpcAft() : pPr.addNewSpcAft();

        if (ctTextSpacing.isSetSpcPct()) {
            ctTextSpacing.unsetSpcPct();
        }

        CTTextSpacingPoint spacing = ctTextSpacing.isSetSpcPts() ? ctTextSpacing.getSpcPts() : ctTextSpacing.addNewSpcPts();
        int pts = (int) (Double.valueOf(pounts) * 100);
        spacing.setVal(pts);
    }

    public void setCTTextParagraphIdent(CTTextParagraph ctTextParagraph, String charsNum) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        pPr.setIndent(Integer.valueOf(charsNum));
    }

    public void test(XSLFTextParagraph paragraph) {
        CTTextParagraphProperties pPr = this.getPPR(paragraph.getXmlObject());
        CTTextSpacing lnSpc = pPr.getLnSpc() == null ? pPr.addNewLnSpc() : pPr.getLnSpc();
        if (lnSpc.isSetSpcPct()) {
            lnSpc.unsetSpcPct();
        }

        CTTextSpacingPoint spcPts = lnSpc.getSpcPts() == null ? lnSpc.addNewSpcPts() : lnSpc.getSpcPts();
        spcPts.setVal(20);
    }


    private String nullToDefault(String goalStr, String defaultStr) {
        if (goalStr == null || "".equals(goalStr)) {
            return defaultStr;
        }
        return goalStr;
    }

    /**
     * 将16进制转换为 byte 数组
     * @param inHex 需要转换的字符串
     * @return
     */
    public byte[] hexToByteArray(String inHex) {
        int hexlen = inHex.length();
        byte[] result;
        if (hexlen % 2 == 1){   // 奇数的话，就在前面添加 0
            hexlen++;
            result = new byte[(hexlen / 2)];
            inHex="0"+inHex;
        }else { // 偶数
            result = new byte[(hexlen / 2)];
        }
        int j=0;
        for (int i = 0; i < hexlen; i += 2){
            result[j] = this.hexToByte(inHex.substring(i, i + 2));
            j++;
        }
        return result;
    }

    private byte hexToByte(String inHex) {
        return (byte) Integer.parseInt(inHex, 16);
    }

}
