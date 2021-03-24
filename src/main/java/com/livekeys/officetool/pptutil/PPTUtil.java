package com.livekeys.officetool.pptutil;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
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
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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

    public XSLFAutoShape getAutoShape(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFAutoShape) {
                return (XSLFAutoShape) shape;
            }
        }
        return null;
    }

    public List<XSLFAutoShape> getAllAutoShape(XSLFSlide slide) {
        List<XSLFAutoShape> autoShapes = new ArrayList<XSLFAutoShape>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFAutoShape) {
                autoShapes.add((XSLFAutoShape) shape);
            }
        }
        return autoShapes;
    }

    // 获取所有幻灯片
    public List<XSLFSlide> getSlides() {
        return pptx.getSlides();
    }

    // 获取所有幻灯片的获取所有图标
    public List<XSLFChart> getCharts() {
        return pptx.getCharts();
    }


    /**
     * 设置幻灯片段落垂直对齐方式
     * @param paragraph
     * @param vertical
     */
    public void setParagraphVerticalAlign(XSLFTextParagraph paragraph, String vertical) {
        vertical = this.nullToDefault(vertical, "auto");

        setCTTextParagraphVerticalAlign(paragraph, vertical.toLowerCase());
    }

    // 设置段落垂直对齐
    private void setCTTextParagraphVerticalAlign(XSLFTextParagraph ctTextParagraph, String verticalStr) {
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

        setCTTextParagraphHorizonAlign(paragraph, horizontal.toLowerCase());
    }

    // 设置段落水平对齐方式
    private void setCTTextParagraphHorizonAlign(XSLFTextParagraph ctTextParagraph, String horizontalStr) {
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
    private CTTextParagraphProperties getPPR(XSLFTextParagraph ctTextParagraph) {
        return ctTextParagraph.getXmlObject().getPPr() == null ? ctTextParagraph.getXmlObject().addNewPPr() : ctTextParagraph.getXmlObject().getPPr();
    }

    /**
     * 设置项目符号的编号
     * @param ctTextParagraph
     * @param lvl
     */
    public void setBulletLvl(XSLFTextParagraph ctTextParagraph, int lvl) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        pPr.setLvl(lvl);
    }

    /**
     * 设置段落自动编号
     * @param ctTextParagraph
     * @param startAt
     */
    public void setAutoBullet(XSLFTextParagraph ctTextParagraph, int startAt) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        CTTextAutonumberBullet bullet = pPr.isSetBuAutoNum() ? pPr.getBuAutoNum() : pPr.addNewBuAutoNum();
        bullet.setStartAt(startAt);
    }

    /**
     * 设置缩进等级，即悬挂缩进
     * @param ctTextParagraph
     * @param indentLevel
     */
    public void setIndentLevel(XSLFTextParagraph ctTextParagraph, int indentLevel) {
        ctTextParagraph.setIndentLevel(indentLevel);
    }

    /**
     * 设置项目符号的颜色
     * @param ctTextParagraph
     * @param colorHex
     */
    public void setBulletColor(XSLFTextParagraph ctTextParagraph, String colorHex) {
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
    public void setLineSpacingPounts(XSLFTextParagraph ctTextParagraph, String pounts) {
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
    public void setLineSpacing(XSLFTextParagraph ctTextParagraph, Double multiple) {

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
    public void setCTTextParagraphSpacingBefore(XSLFTextParagraph ctTextParagraph, String pounts) {
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
    public void setCTTextParagraphSpacingAfter(XSLFTextParagraph ctTextParagraph, String pounts) {
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

    /**
     * 设置缩进字符
     * @param ctTextParagraph
     * @param charsNum
     */
    public void setCTTextParagraphIndent(XSLFTextParagraph ctTextParagraph, String charsNum) {
        CTTextParagraphProperties pPr = this.getPPR(ctTextParagraph);
        pPr.setIndent(Integer.valueOf(charsNum));
    }

    /**
     * 为一个段落添加文本，appendText 参数为 true 的话，就追加文本，false 的话就重开头设置文本
     * @param paragraph  要操作的段落
     * @param text  文本
     * @param appendText    是否追加文本
     */
    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph, Boolean appendText, String text, Boolean bold) {
        XSLFTextRun textRun = getNewRun(paragraph, appendText);
        textRun.setText(text);
        textRun.setBold(bold);
        return textRun;
    }

    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph,
                                        Boolean appendText,
                                        String text,
                                        Boolean bold,
                                        String chinesefontFamily,
                                        String westernFontFamily,
                                        String fontSize) {
        XSLFTextRun textRun = addParagraphText(paragraph, appendText, text, bold);    // 添加文本
        textRun.setFontSize(Double.valueOf(fontSize));  // 设置字体大小

        CTTextCharacterProperties rPr = getRPr(textRun.getXmlObject());
        setRPRFontFamily(rPr, chinesefontFamily, westernFontFamily);    // 设置字体

        return textRun;
    }

    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph,
                                        Boolean appendText,
                                        String text,
                                        Boolean bold,
                                        String chinesefontFamily,
                                        String westernFontFamily,
                                        String fontSize,
                                        String color) {
        XSLFTextRun textRun = addParagraphText(paragraph, appendText, text, bold, chinesefontFamily, westernFontFamily, fontSize);    // 添加文本
        textRun.setFontSize(Double.valueOf(fontSize));  // 设置字体大小
        CTTextCharacterProperties rPr = getRPr(textRun.getXmlObject());

        // 设置字体颜色
        CTSolidColorFillProperties solidColor = rPr.isSetSolidFill() ? rPr.getSolidFill() : rPr.addNewSolidFill();
        CTSRgbColor ctColor = solidColor.isSetSrgbClr() ? solidColor.getSrgbClr() : solidColor.addNewSrgbClr();
        ctColor.setVal(hexToByteArray(color.substring(1)));
        return textRun;
    }

    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph,
                                        Boolean appendText,
                                        String text,
                                        Boolean bold,
                                        String chinesefontFamily,
                                        String westernFontFamily,
                                        String fontSize,
                                        String color,
                                        Boolean italic) {
        XSLFTextRun textRun = addParagraphText(paragraph, appendText, text, bold, chinesefontFamily, westernFontFamily, fontSize, color);    // 添加文本
        textRun.setItalic(italic);
        return textRun;
    }

    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph,
                                        Boolean appendText,
                                        String text,
                                        Boolean bold,
                                        String chinesefontFamily,
                                        String westernFontFamily,
                                        String fontSize,
                                        String color,
                                        Boolean italic,
                                        Boolean strike) {
        XSLFTextRun textRun = addParagraphText(paragraph, appendText, text, bold, chinesefontFamily, westernFontFamily, fontSize, color, italic);    // 添加文本
        textRun.setStrikethrough(strike);
        return textRun;
    }

    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph,
                                        Boolean appendText,
                                        String text,
                                        Boolean bold,
                                        String chinesefontFamily,
                                        String westernFontFamily,
                                        String fontSize,
                                        String color,
                                        Boolean italic,
                                        Boolean strike,
                                        Boolean underline) {
        XSLFTextRun textRun = addParagraphText(paragraph, appendText, text, bold, chinesefontFamily, westernFontFamily, fontSize, color, italic, strike);    // 添加文本
        textRun.setUnderlined(underline);
        return textRun;
    }

    public XSLFTextRun addParagraphText(XSLFTextParagraph paragraph,
                                        Boolean appendText,
                                        String text,
                                        Boolean bold,
                                        String chinesefontFamily,
                                        String westernFontFamily,
                                        String fontSize,
                                        String color,
                                        Boolean italic,
                                        Boolean strike,
                                        Boolean underline,
                                        String characterSpacing) {
        XSLFTextRun textRun = addParagraphText(paragraph, appendText, text, bold, chinesefontFamily, westernFontFamily, fontSize, color, italic, strike, underline);    // 添加文本
        textRun.setCharacterSpacing(Double.valueOf(characterSpacing));
        return textRun;
    }

    // 替换段内的标签文本
    public void replaceTagInParagraph(XSLFTextParagraph paragraph, Map<String, Object> paramMap) {

        String paraText = paragraph.getText();
        String regEx = "\\{.+?\\}";
        Pattern pattern = Pattern.compile(regEx);
        Matcher matcher = pattern.matcher(paraText);

        if (matcher.find()) {
            StringBuilder keyWord = new StringBuilder();
            int start = getRunIndex(paragraph, "{");
            int end = getRunIndex(paragraph, "}");

            // 处理 {***} 在一个 run 内的情况
            if (start == end) {
                String rs = matcher.group(0);   // 存放标签
                keyWord.append(rs.replace("{", "").replace("}", ""));   // 存放 key
                String text = getRunsT(paragraph, start, end + 1);
                String v = nullToDefault(paramMap.get(keyWord.toString()), keyWord.toString());
                setText(paragraph.getTextRuns().get(start), text.replace(rs, v));

            }

            replaceTagInParagraph(paragraph, paramMap); // 继续找
        }

    }

    private String getRunsT(XSLFTextParagraph paragraph, int start, int end) {
        List<XSLFTextRun> textRuns = paragraph.getTextRuns();
        StringBuilder t = new StringBuilder();
        for (int i = start; i < end; i++) {
            t.append(textRuns.get(i).getRawText());
        }
        return t.toString();
    }

    private void setText(XSLFTextRun run, String t) {
        run.setText(t);
    }

    private int getRunIndex(XSLFTextParagraph paragraph, String word) {
        List<CTRegularTextRun> rList = paragraph.getXmlObject().getRList();
        for (int i = 0; i < rList.size(); i++) {

            String text = rList.get(i).getT();
            if (text.contains(word)) {
                return i;
            }
        }
        return -1;
    }

    // 设置 rPr 的字体
    private void setRPRFontFamily(CTTextCharacterProperties rPr, String chinesefontFamily, String westernFontFamily) {
        if (chinesefontFamily == null || "".equals(chinesefontFamily)) {
            chinesefontFamily = "宋体";
        }

        if (westernFontFamily == null || "".equals(westernFontFamily)) {
            westernFontFamily = "宋体";
        }


        if (rPr.isSetLatin()) {
            rPr.unsetLatin();
        }

        CTTextFont ea = rPr.isSetEa() ? rPr.getEa() : rPr.addNewEa();
        ea.setTypeface(chinesefontFamily);
        ea.setPitchFamily(new Byte("34"));
        ea.setCharset(new Byte("-122"));

        CTTextFont cs = rPr.isSetCs() ? rPr.getCs() : rPr.addNewCs();
        cs.setTypeface(chinesefontFamily);
        cs.setPitchFamily(new Byte("34"));
        cs.setCharset(new Byte("-122"));

        CTTextFont latin = rPr.isSetLatin() ? rPr.getLatin() : rPr.addNewLatin();
        latin.setTypeface(westernFontFamily);
        latin.setPitchFamily(new Byte("34"));
        latin.setCharset(new Byte("-122"));
    }

    // 获取新添加的 run
    private XSLFTextRun getNewRun(XSLFTextParagraph paragraph, Boolean appendText)  {
        if (!appendText) {  // 是否追加文本
            this.clearParagraphText(paragraph);
        }

        return paragraph.addNewTextRun();
    }

    private void clearParagraphText(XSLFTextParagraph paragraph) {
        CTTextParagraph ctTextParagraph = paragraph.getXmlObject();
//        int s = ctTextParagraph.getRList().size();
//        for (int i = 0; i < s; i++) {
//            ctTextParagraph.removeR(0);
//        }
        ctTextParagraph.getRList().clear();
        paragraph.getTextRuns().clear();
    }

    // 获取 rPR
    private CTTextCharacterProperties getRPr(XmlObject xmlObject) {
        if (xmlObject instanceof CTTextField) {
            CTTextField tf = (CTTextField) xmlObject;
            return tf.getRPr() == null ? tf.addNewRPr() : tf.getRPr();
        } else if (xmlObject instanceof CTTextLineBreak) {
            CTTextLineBreak tlb = (CTTextLineBreak) xmlObject;
            return tlb.getRPr() == null ? tlb.addNewRPr() : tlb.getRPr();
        } else {
            CTRegularTextRun tr = (CTRegularTextRun) xmlObject;
            return tr.getRPr() == null ? tr.addNewRPr() : tr.getRPr();
        }
    }



   // 空字符串转默认值
    private String nullToDefault(String goalStr, String defaultStr) {
        if (goalStr == null || "".equals(goalStr)) {
            return defaultStr;
        }
        return goalStr;
    }

    private String nullToDefault(Object o, String defaultStr) {
        if (o == null) {
            return defaultStr;
        } else {
            if ("".equals(o.toString())) {
                return defaultStr;
            } else {
                return o.toString();
            }
        }
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
