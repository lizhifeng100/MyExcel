package com.greeoneplus.poi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/*
 *author liZhiFeng
 * work in gree 2020/22/14
 * 这个类的主要作用是设置格式，比如：表头、字段、值的格式
 * 然后通过ExcelWriter调用，设置相关格式。
 */
public class SheetStyle {

    private CellStyle cellStyle;

    private Workbook workbook;


    public SheetStyle(Workbook workbook) {
        this.workbook = workbook;
        this.cellStyle = workbook.createCellStyle();
    }

    public SheetStyle() {

    }


    /*
     *设置单元格字体的格式
     * 字体一般有什么格式？大小、样式、颜色、加粗、
     */
    public Font setFront() {

        Font font = this.workbook.createFont();


        return null;

    }

}
