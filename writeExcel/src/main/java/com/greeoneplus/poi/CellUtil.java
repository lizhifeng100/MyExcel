package com.greeoneplus.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.concurrent.atomic.AtomicInteger;

/*
 *写这个类的原因是：主要是为了进行单元格的操作，
 * 比如：合并单元格、将格式（sheetStyle）写入到单元格中，
 *
 *
 *
 */
public class CellUtil {

    private AtomicInteger currentRow = new AtomicInteger(0);

    /*还没有测试？ok：no
     * 合并单元格，
     *
     * value:合并单元格后，里面的值。
     */
    public void merge(Sheet sheet, Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol, String value) {

        Row row = sheet.createRow(firstRow);
        Cell cell = row.createCell(firstCol);
        cell.setCellValue(value);
        CellRangeAddress cellAddresses = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

        sheet.addMergedRegion(cellAddresses);

        this.currentRow.incrementAndGet();


    }
}
