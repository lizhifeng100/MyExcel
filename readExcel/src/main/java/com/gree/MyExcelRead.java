package com.gree;

import com.sun.org.slf4j.internal.Logger;
import com.sun.org.slf4j.internal.LoggerFactory;
import jdk.internal.org.xml.sax.InputSource;
import jdk.internal.org.xml.sax.XMLReader;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

//数据量大的时候采取Sax模式来读Excel文件
//还有版本问题，03版使用的是“.xls”，，07版使用的是“.xlsx”，同样存储数据多少也有不同。
//  03版使用的是：HSSFWorkbook，07版使用的是：XSSFWorkbook，虽然可以读取104万行，但是回oom，使用SXSSFWorkbook可以避免这个问题。
//总的来说，根据数据量的大小来判断用哪种方式来读取Excel。

public class MyExcelRead {

    public static final Logger logger = LoggerFactory.getLogger(MyExcelRead.class);
    private static final Object T = null;

    public static StylesTable stylesTable;

    public void processOneSheet(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        stylesTable = r.getStylesTable();
//        XMLReader parser = fetchSheetParser(sst);

        // To look up the Sheet Name / Sheet Order / rID,
        // you need to process the core Workbook stream.
        // Normally it's of the form rId# or rSheet#
        InputStream sheet = r.getSheet("rId1");
        InputSource sheetSource = new InputSource(sheet);
//        parser.parse(sheetSource);
        sheet.close();
    }

    private POIFSFileSystem fs;
    private HSSFWorkbook wb;
    private HSSFSheet sheet;
    private HSSFRow row;

    //读取表头title
    public String[] readExcelTitle(InputStream is) {
        try {
            fs = new POIFSFileSystem(is);
            wb = new HSSFWorkbook(fs);
        } catch (IOException e) {
            e.printStackTrace();
        }
        sheet = wb.getSheetAt(0);
        row = sheet.getRow(0);
        // 标题总列数
        int colNum = row.getPhysicalNumberOfCells();
        System.out.println("colNum:" + colNum);
        String[] title = new String[colNum];
        for (int i = 0; i < colNum; i++) {
            //title[i] = getStringCellValue(row.getCell((short) i));
            title[i] = getCellFormatValue(row.getCell((short) i));
        }
        return title;
    }


    //读取Excel表格的内容,读取数据放在Map中，Integer代表行，String代表
    private void readExcelContent(Map<Integer, String> content, Integer rowNum, Integer colNum) {
        content = new HashMap<Integer, String>();
        StringBuffer str = new StringBuffer("");
        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            int j = 0;
            while (j < colNum) {
                str.append(getCellFormatValue(row.getCell((short) j)).trim() + "_");
                j++;
            }
            content.put(i, str.toString());
            //清空StringBuffer
            str.setLength(0);
        }
    }

    //读取Excel表格的内容,读取数据放在Map中，Integer代表行，List<T> 代表将数据放在写好的javaBean中。
    public void readExcelContent(InputStream is, Boolean isJavaBean, Map<Integer, T> map) {
        try {
            fs = new POIFSFileSystem(is);
            wb = new HSSFWorkbook(fs);
        } catch (IOException e) {
            e.printStackTrace();
        }
        this.sheet = wb.getSheetAt(0);
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        logger.debug(rowNum + "**************");

        this.row = sheet.getRow(0);
        //获得一行中列数
        int colNum = row.getPhysicalNumberOfCells();
        if (!isJavaBean) {
            if (T instanceof String) {
                 new HashMap<Integer, String>();
            }
//            readExcelContent(, rowNum, colNum);
        } else {

        }
        Map<Integer, List<T>> content = new HashMap<Integer, List<T>>();
        List<T> tList = new ArrayList<T>();

        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            int j = 0;
            while (j < colNum) {
                // 每个单元格的数据内容用"-"分割开，以后需要时用String类的replace()方法还原数据
                // 也可以将每个单元格的数据设置到一个javabean的属性中，此时需要新建一个javabean
                // str += getStringCellValue(row.getCell((short) j)).trim() +
                // "-";

                String cellFormatValue = getCellFormatValue(row.getCell(j));
//                str.append(getCellFormatValue(row.getCell((short) j)).trim());

                j++;
            }
//            content.put(i, str.toString());
            //
//            str.setLength(0);
        }
    }

    //读取HSSFCell表格中的数据,比如数字、日期、公式、字符串。
    private String getCellFormatValue(HSSFCell cell) {
        String cellValue = "";
        if (cell != null) {
            int code = cell.getCellType().getCode();
            // 判断当前Cell的Type
            switch (code) {
                // 如果当前Cell的Type为NUMERIC
                case 0:         //Numeric数字
                case 2: {       //Formula公式
                    // 判断当前的cell是否为Date
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        // 如果是Date类型则，转化为Data格式

                        //方法1：这样子的data格式是带时分秒的：2011-10-12 0:00:00
                        cellValue = cell.getDateCellValue().toLocaleString();

                        //方法2：这样子的data格式是不带带时分秒的：2011-10-12
//                        Date date = cell.getDateCellValue();
//                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
//                        cellvalue = sdf.format(date);

                    }
                    // 如果是纯数字
                    else {
                        // 取得当前Cell的数值
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case 1:
                    // 取得当前的Cell字符串
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                case 4:
                    //Boolean值
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                // 默认的Cell值
                default:
                    //code：3是blank
                    //code：5是error值
                    cellValue = " ";
            }
        } else {
            cellValue = " ";
        }
        return cellValue;

    }
}
