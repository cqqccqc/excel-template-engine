package com.ctrip.mice.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.regex.Pattern;

/**
 * Created by q.chen on 2016/9/23.
 */
public class ExcelTemplateEngine {

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

    private HSSFWorkbook workbook;

    /**
     * constructor
     * initialize the excel template
     * @param filePath file path
     * @throws IOException if the file not found
     */
    public ExcelTemplateEngine(String filePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(filePath);
        workbook = new HSSFWorkbook(fileInputStream);
    }

    /**
     * constructor
     * initialize the excel template
     * @param inputStream file input stream
     * @throws IOException if the stream cannot be read
     */
    public ExcelTemplateEngine(InputStream inputStream) throws IOException {
        workbook = new HSSFWorkbook(inputStream);
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
     * write to output stream
     * @return a new instance of output stream
     * @throws IOException if cannot write
     */
    public OutputStream writeToStream() throws IOException{
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream;
    }

}
