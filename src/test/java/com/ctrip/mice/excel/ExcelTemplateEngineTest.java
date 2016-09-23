package com.ctrip.mice.excel;

import org.junit.Test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.junit.Assert.*;

/**
 * Created by q.chen on 2016/9/23.
 */
public class ExcelTemplateEngineTest {

    @Test
    public void testBrace() throws Exception {
        Pattern loopText = Pattern.compile("\\{loop:([a-zA-Z_0-9]+):([a-zA-Z_0-9.]+)}");
        Matcher m = loopText.matcher("aaa{loop:sub:List}bbb");
        StringBuffer sb = new StringBuffer(100);
        int i = 0;
        //m.appendTail(sb);
        System.out.println(m.group());
    }

}