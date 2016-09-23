package com.ctrip.mice.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * Created by q.chen on 2016/9/23.
 * Excel Template Engine
 */
public class ExcelTemplateEngine<T> {

    /**
     * to match loop
     */
    private Pattern loopText = Pattern.compile("\\{loop:([a-zA-Z_0-9]+):([a-zA-Z_0-9.]+)}");
    /**
     * to match variable text
     */
    private Pattern varNameText = Pattern.compile("\\{([a-zA-Z_0-9]+)}");
    /**
     * to match all {xxx}
     */
    private Pattern matchAllText = Pattern.compile("\\{([a-zA-Z_0-9:.#]+)}");
    /**
     * to match a line to be deleted.
     */
    private Pattern toBeDelText = Pattern.compile("\\{(#ToBeDeleted#)}");
    /**
     * to if condition
     */
    private Pattern ifText = Pattern.compile("\\{if:([a-zA-Z_0-9]+):([a-zA-Z_0-9]+)}");
    /**
     * to include statement {include:templatename:varname}
     */
    private Pattern includeText = Pattern.compile("\\{include:([a-zA-Z_0-9]+):([a-zA-Z_0-9]+)}");

    /**
     * excel file instance
     */
    private XSSFWorkbook workbook;

    private Map<String, T> dataSource;

    /**
     * constructor
     * initialize the excel template
     * @param filePath file path
     * @throws IOException if the file not found
     */
    public ExcelTemplateEngine(String filePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(filePath);
        workbook = new XSSFWorkbook(fileInputStream);
    }

    /**
     * constructor
     * initialize the excel template
     * @param inputStream file input stream
     * @throws IOException if the stream cannot be read
     */
    public ExcelTemplateEngine(InputStream inputStream) throws IOException {
        workbook = new XSSFWorkbook(inputStream);
    }

    /**
     * write to output stream
     * @param outputStream out put stream to write
     * @throws IOException if cannot write
     */
    public void writeToStream(ByteArrayOutputStream outputStream) throws IOException{
        workbook.write(outputStream);
    }

    /**
     * get result output stream
     * @return result output stream
     * @throws IOException if cannot write
     */
    public OutputStream getResultOutputStream() throws IOException{
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream;
    }

    /**
     * Render Engine
     * @param mainTemplateName excel name
     * @param dataToRender data that to be render
     */
    public void Render(String mainTemplateName, T dataToRender) {
        XSSFSheet wsMain = workbook.getSheet(mainTemplateName);
        int rowEnd = wsMain.getLastRowNum();
        int colEnd = getLastColNum(wsMain);
    }

    /**
     * count max column number
     * @param sheet work sheet
     * @return last column number
     */
    private int getLastColNum(XSSFSheet sheet) {
        int maxCol = 1;
        Iterator<Row> iterator =sheet.rowIterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            int lastCellNum = row.getLastCellNum();
            maxCol = lastCellNum > maxCol ? lastCellNum : maxCol;
        }
        return maxCol;
    }


    private void RenderPrimitiveValue(){

    }

}
