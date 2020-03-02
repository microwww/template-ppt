package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class AppendOperationTest extends RepeatOperationTest {

    @Test
    public void append() throws IOException {
        ParseContext context = super.createContext(0);
        XSLFShape shape = context.getTemplate().getShapes().get(0);
        XSLFTable table = (XSLFTable) shape;
        String express = new StringBuffer().append("append XSLFTable 0 XSLFTableRow 1 ( 'append-test' null 'ok' )").toString();
        express += "\n" + express;
        ByteArrayInputStream input = new ByteArrayInputStream(express.getBytes());
        //context.parse(input, new FileOutputStream("C:\\Users\\changshu.li\\Desktop\\test.pptx"));
    }
}