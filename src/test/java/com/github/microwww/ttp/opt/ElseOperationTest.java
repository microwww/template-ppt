package com.github.microwww.ttp.opt;

import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

import static org.junit.Assert.*;

public class ElseOperationTest {

    @Test
    public void testOpt() throws IOException {
        ParseContext context = new RepeatOperationTest().createContext(1);
        IfOperation ifo = new IfOperation();
        ifo.setParams(new String[]{"1", "+", "1", "==2"});
        {
            ReplaceOperation rep = new ReplaceOperation();
            rep.setNode(new String[]{"XSLFTable", "0", "XSLFTableRow", "0", "XSLFTableCell", "0"});
            rep.setParams(new String[]{"'name'"});
            ifo.addChildrenOperation(rep);
        }
        {
            IfOperation ifo2 = new IfOperation();
            ifo.addChildrenOperation(ifo2);
            ifo2.setParams(new String[]{"1", "!=", "1"});
            {
                ReplaceOperation rep = new ReplaceOperation();
                rep.setNode(new String[]{"XSLFTable", "0", "XSLFTableRow", "2", "XSLFTableCell", "0"});
                rep.setParams(new String[]{"'Good'"});
                ifo2.addChildrenOperation(rep);
            }
            {
                ElseOperation rep = new ElseOperation();
                ifo2.addChildrenOperation(rep);
            }
            {
                ReplaceOperation rep = new ReplaceOperation();
                rep.setNode(new String[]{"XSLFTable", "0", "XSLFTableRow", "3", "XSLFTableCell", "0"});
                rep.setParams(new String[]{"'NoGood'"});
                ifo2.addChildrenOperation(rep);
            }
        }
        ifo.parse(context);
        context.write(new FileOutputStream(FileUtils.getFile(FileUtils.getUserDirectory(), "Desktop", "test.pptx")));
    }

}