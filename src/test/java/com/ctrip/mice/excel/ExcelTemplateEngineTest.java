package com.ctrip.mice.excel;

import org.junit.Test;

import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.junit.Assert.*;

/**
 * Created by q.chen on 2016/9/23.
 * Test Excel Template Engine
 */
public class ExcelTemplateEngineTest {

    @Test
    public void testBraceRegex() throws Exception {
        Pattern loopText = Pattern.compile("\\{loop:([a-zA-Z_0-9]+):([a-zA-Z_0-9.]+)}");
        Matcher m = loopText.matcher("aaa{loop:sub:List}bbb");
        System.out.println(m.group());
    }

    @Test
    public void testExcelTemplateEngine() {
        try {
            ClassLoader classLoader = getClass().getClassLoader();
            String path = classLoader.getResource("OrderListReport.xlsx").getPath();
            ExcelTemplateEngine<OrderList> excelTemplateEngine = new ExcelTemplateEngine<>(path);
            System.out.print(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}

class OrderList {

}