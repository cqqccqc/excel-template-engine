package com.ctrip.mice.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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

    private final String filePath = "D:\\test.xlsx";

    // private final String filePathForMac = "/Users/chenqi/Develop/test.xlsx";

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
            ExcelTemplateEngine excelTemplateEngine = new ExcelTemplateEngine(path);
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
        System.out.println(matcherTrue.group());
        System.out.println(canFind);
        Assert.assertTrue(canFind);

        Matcher matcherFalse = matchAllText.matcher("fsdfsfd");
        boolean cantFind = matcherFalse.find();
        System.out.println(cantFind);
        Assert.assertFalse(cantFind);

        Pattern varNameText = Pattern.compile("\\{([a-zA-Z_0-9]+)}");
    }

    /**
     * Test render primitive value
     * @throws IOException can't write file
     * @throws NullPointerException  can't get file
     */
    @Test
    public void testRenderPrimitive() throws IOException, NullPointerException {
        ExcelTemplateEngine templateEngine =
                new ExcelTemplateEngine(getClass().getClassLoader().getResource("OrderTest.xlsx").getPath());
        templateEngine.renderTemplate(templateEngine.workbook.getSheet("订单"), new Order(), 0, 0, 1, 3);
        File file = new File(filePath);
        if(!file.exists()) {
            boolean created = file.createNewFile();
            Assert.assertTrue(created);
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);

        ByteArrayOutputStream byteArrayOutputStream = (ByteArrayOutputStream)templateEngine.getResultOutputStream();
        fileOutputStream.write(byteArrayOutputStream.toByteArray());
    }

    @Test
    public void testRenderPrimitiveValue() throws IOException {
        ExcelTemplateEngine templateEngine =
                new ExcelTemplateEngine(getClass().getClassLoader().getResource("OrderTest.xlsx").getPath());
        Sheet sheet = templateEngine.workbook.getSheet("订单");
        templateEngine.renderPrimitiveValue(sheet.getRow(1).getCell(1), null);

        File file = new File(filePath);
        if(!file.exists()) {
            boolean created = file.createNewFile();
            Assert.assertTrue(created);
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);

        ByteArrayOutputStream byteArrayOutputStream = (ByteArrayOutputStream)templateEngine.getResultOutputStream();
        fileOutputStream.write(byteArrayOutputStream.toByteArray());
    }

    public void testMatchInclude() throws IOException {
        ExcelTemplateEngine templateEngine =
                new ExcelTemplateEngine(getClass().getClassLoader().getResource("OrderTest.xlsx").getPath());
    }
}

