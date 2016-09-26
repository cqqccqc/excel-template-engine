package com.ctrip.mice.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by q.chen on 2016/9/23.
 * Excel Template Engine
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

    /**
     * excel file instance
     */
    public XSSFWorkbook workbook;

    private Map<String, Object> dataSource;

    /**
     * constructor
     * initialize the excel template
     *
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
     *
     * @param inputStream file input stream
     * @throws IOException if the stream cannot be read
     */
    public ExcelTemplateEngine(InputStream inputStream) throws IOException {
        workbook = new XSSFWorkbook(inputStream);
    }

    /**
     * write to output stream
     *
     * @param outputStream out put stream to write
     * @throws IOException if cannot write
     */
    public void writeToStream(ByteArrayOutputStream outputStream) throws IOException {
        workbook.write(outputStream);
    }

    /**
     * get result output stream
     *
     * @return result output stream
     * @throws IOException if cannot write
     */
    public OutputStream getResultOutputStream() throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream;
    }

    /**
     * count max column number
     *
     * @param sheet work sheet
     * @return last column number
     */
    private int getLastColNum(XSSFSheet sheet) {
        int maxCol = 1;
        Iterator<Row> iterator = sheet.rowIterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            int lastCellNum = row.getLastCellNum();
            maxCol = lastCellNum > maxCol ? lastCellNum : maxCol;
        }
        return maxCol;
    }

    /**
     * Render Engine
     *
     * @param mainTemplateName excel name
     * @param dataToRender     data that to be render
     */
    public void render(String mainTemplateName, Object dataToRender) {
        XSSFSheet wsMain = workbook.getSheet(mainTemplateName);
        // this value will be calculated and updated after insert value into the sheet
        int rowEnd = wsMain.getLastRowNum();
        int colEnd = getLastColNum(wsMain);
    }

    /**
     * Render function
     * called recursively to render every entry
     * if this cell's value is template string, then render next cell.
     * if this cell's value is field name, render field's value directly.
     * otherwise if this cell's value is a template direction,
     * check if it is a 'loop' direction or not.
     * if it is an 'if' direction, evaluate its value and parse it into 'include' direction.
     * Then render the sub template by call the render function recursively.
     * Finally, if it is a 'loop' direction, call the render function recursively throw each entry of the list
     *
     * @param sheet    the sheet to be rendered
     * @param rowStart row start
     * @param colStart column start
     * @param rowEnd   row end
     * @param colEnd   column end
     */
    public void renderTemplate(XSSFSheet sheet, Object dataToRender, int rowStart, int colStart, int rowEnd, int colEnd) {
        // if rowStart > rowEnd and colStart > colEnd, end render
        if (rowStart > rowEnd && colStart > colEnd) return;

        // get row
        Row row = sheet.getRow(rowStart);
        if (row == null) {
            return;
        }
        // get cell
        Cell cell = sheet.getRow(rowStart).getCell(colStart);
        if (cell == null) {
            renderTemplate(sheet, dataToRender, ++rowStart, 0, rowEnd, colEnd);
            return;
        }

        String value = cell.getStringCellValue();
        // if value not match '{}', which means it is just a normal template string, continue to render next cell
        Matcher matcher = matchAllText.matcher(value);
        if (!matcher.find())
            renderTemplate(sheet, dataToRender, rowStart, ++colStart, rowEnd, colEnd);
        else if (renderPrimitiveValue(cell, dataToRender)) {
            renderTemplate(sheet, dataToRender, rowStart, ++colStart, rowEnd, colEnd);
        }
    }

    /**
     * render primitive value
     * @param cell excel cell
     * @param dataToRender data to render
     * @return whether the cell value match the primitive value syntax {variable}
     */
    public boolean renderPrimitiveValue(Cell cell, Object dataToRender) {
        Matcher matcher = varNameText.matcher(cell.getStringCellValue());
        if (matcher.find()) {
            String matched = matcher.group();
            matched = matched.substring(1, matched.length() - 1);
            try {
                // null object to render an empty string
                if (dataToRender == null) {
                    cell.setCellValue("");
                    return true;
                }
                Field field = dataToRender.getClass().getDeclaredField(matched);
                field.setAccessible(true);
                Object replace = field.get(dataToRender);
                cell.setCellValue(replace.toString());
            } catch (IllegalAccessException | NoSuchFieldException e) {
                e.printStackTrace();
                return false;
            }
            return true;
        }
        return false;
    }

    /**
     * render {include:template name:var name}
     * @param cell excel cell
     * @param template template name
     * @param dataToRender data to render
     * @return whether the cell match template syntax {include:template name:var name}
     */
    public boolean renderIncludeTemplate(Cell cell, String template, Object dataToRender) {
        Matcher matcher = includeText.matcher(cell.getStringCellValue());
        if(matcher.find()) {
            String matched = matcher.group();
            return true;
        }
        return false;
    }
}
