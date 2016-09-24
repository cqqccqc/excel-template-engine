package com.ctrip.mice.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.*;
import java.util.Iterator;
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
            ExcelTemplateEngine<Order> excelTemplateEngine = new ExcelTemplateEngine<>(path);
            System.out.print(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private XSSFWorkbook createWorkbook(String fileName) throws IOException {
        ClassLoader classLoader = getClass().getClassLoader();
        String path = classLoader.getResource(fileName).getPath();
        return new XSSFWorkbook(new FileInputStream(path));
    }



    @Test
    public void testExcelIterator() {
        try {
            String fileName = "OrderTest.xlsx";
            XSSFWorkbook workbook = createWorkbook(fileName);
            XSSFSheet sheet = workbook.getSheet("订单列表");
            Iterator<Row> rowIterator = sheet.rowIterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    System.out.println(cell.getStringCellValue());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testMatcher(){
        Pattern matchAllText = Pattern.compile("\\{([a-zA-Z_0-9:.#]+)}");
        Matcher matcherTrue = matchAllText.matcher("fds{fff}fsd");
        boolean canFind = matcherTrue.find();
        System.out.println(canFind);
        Assert.assertTrue(canFind);

        Matcher matcherFalse = matchAllText.matcher("fsdfsfd");
        boolean cantFind = matcherFalse.find();
        System.out.println(cantFind);
        Assert.assertFalse(cantFind);
    }

    @Test
    public void testRender() throws IOException, NullPointerException {
        ExcelTemplateEngine<Order> templateEngine =
                new ExcelTemplateEngine<Order>(getClass().getClassLoader().getResource("OrderTest.xlsx").getPath());
        templateEngine.RenderTemplate(templateEngine.workbook.getSheet("订单"), 1, 1, 2, 4);
        File file = new File("/Users/chenqi/Develop/test.xlsx");
        if(!file.exists()) {
            file.createNewFile();
        }
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/chenqi/Develop/test.xlsx");

        ByteArrayOutputStream byteArrayOutputStream = (ByteArrayOutputStream)templateEngine.getResultOutputStream();
        fileOutputStream.write(byteArrayOutputStream.toByteArray());
    }
}

