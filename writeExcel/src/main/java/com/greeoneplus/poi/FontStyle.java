package com.greeoneplus.poi;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;

/*
 *author liZhiFeng
 * work in gree 2020/22/14
 * 这个类的主要作用是设置格式，比如：表头、字段、值的格式
 * 然后通过ExcelWriter调用，设置相关格式。
 */
public class FontStyle {

    private CellStyle cellStyle;

    private Workbook workbook;


    public FontStyle(Workbook workbook) {
        this.workbook = workbook;
        this.cellStyle = workbook.createCellStyle();
    }


    /*还没有测试？ok：no
     *设置单元格字体的格式
     * 字体一般有什么格式？大小、样式、颜色、加粗、
     * 我打算用一个map将所有的字体样式包含进去，
     * ************************************************
     * 此方法应该有改进的空间，我只是用if……else if来判断所有的情况
     */
    public void setFront(Map<String, String> fontStyleMap) {

        Font font = this.workbook.createFont();
        font.setCharSet(HSSFFont.DEFAULT_CHARSET);
        for (String key : fontStyleMap.keySet()) {

            if (key.equals("font-bold")) { //字体是否加粗
                String boldValue = fontStyleMap.get(key);
                font.setBold(Boolean.valueOf(boldValue));
            } else if (key.equals("font-size")) { //字体大小
                String sizeValue = fontStyleMap.get(key);
                font.setFontHeightInPoints(Short.valueOf(sizeValue));
            } else if (key.equals("font-color")) { //字体颜色
                String colorValue = fontStyleMap.get(key);
                font.setColor(Short.valueOf(colorValue));
            } else if (key.equals("font-style")) { //字体样式：比如“宋体”
                String style = fontStyleMap.get(key);
                font.setFontName(style);
            } else if (key.equals("font-italic")) { // 字体是否斜体
                String slantedValue = fontStyleMap.get(key);
                font.setItalic(Boolean.valueOf(slantedValue));
            } else if (key.equals("font-underLine")) { //下划线
                String underLineValue = fontStyleMap.get(key);
                if (null != underLineValue) {
                    if (underLineValue.equals("U_NONE")) {
                        font.setUnderline(Font.U_NONE);
                    } else if (underLineValue.equals("U_SINGLE")) {
                        font.setUnderline(Font.U_SINGLE);
                    } else if (underLineValue.equals("U_DOUBLE")) {
                        font.setUnderline(Font.U_DOUBLE);
                    } else if (underLineValue.equals("U_SINGLE_ACCOUNTING")) {
                        font.setUnderline(Font.U_SINGLE_ACCOUNTING);
                    } else if (underLineValue.equals("U_DOUBLE_ACCOUNTING")) {
                        font.setUnderline(Font.U_DOUBLE_ACCOUNTING);
                    }
                }
            } else if (key.equals("font-center-horizontally")) { // 水平居中
                String centerHorizontally = fontStyleMap.get(key);
                if (Boolean.valueOf(centerHorizontally)) {
                    this.cellStyle.setAlignment(HorizontalAlignment.CENTER);
                }
            } else if (key.equals("font-center-vertically")) { // 垂直居中
                String centerVertically = fontStyleMap.get(key);
                if (Boolean.valueOf(centerVertically)) {
                    this.cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                }
            }

        }

        cellStyle.setWrapText(true);
        this.cellStyle.setFont(font);
    }




}
