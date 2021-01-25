import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import org.junit.Test;

import java.awt.datatransfer.Clipboard;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MyTest {


    @Test
    public void testWrite() {
        List<String> row1 = CollUtil.newArrayList("aa", "bb", "cc", "dd");
        List<String> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1");
        List<String> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2");
        List<String> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3");
        List<String> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4");

        List<List<String>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);
        ExcelWriter writer = ExcelUtil.getWriter("d:/test.xlsx");
        Map<Integer, String> map = new HashMap<Integer, String>();
        List<Map<Integer, String>> mapArrayList = new ArrayList<Map<Integer, String>>();
        writer.write(mapArrayList,true);
        writer.write(rows, true);
        writer.close();


    }

    @Test
    public void testRead() {
//        Excel07SaxReader reader = new Excel07SaxReader(createRowHandler());
//        reader.read("d:/test.xlsx", 0);

        ExcelReader reader = ExcelUtil.getReader("d:/aaa.xlsx");
        List<List<Object>> readAll = reader.read();
    }


    private RowHandler createRowHandler() {
        return new RowHandler() {
            public void handle(int sheetIndex, long rowIndex, List<Object> rowlist) {
                Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowlist);
                System.out.println("****");
                System.out.println(sheetIndex);
                System.out.println(rowIndex);
                for (Object o : rowlist) {
                    System.out.println(o);
                }
            }
        };
    }

}
