package com.gree;

import exception.MyException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.Closeable;
import java.io.IOException;

//这个类有什么作用？自己看哦，因为这是我自己的工具类。
//虽然我尽量做到规范，但是我不想写这个注释，因为今天星期天
//author:lzf/260494
public class ExcelReader implements Closeable {

    //是否关闭
    protected boolean isClosed;

    //工作薄
    protected Workbook workbook;

    //Excel中的sheet
    protected Sheet sheet;

    public ExcelReader(Sheet sheet) {
        MyException.notNull(sheet, "sheet为空！");
        this.sheet = sheet;
    }

    public void close() throws IOException {

    }
}
