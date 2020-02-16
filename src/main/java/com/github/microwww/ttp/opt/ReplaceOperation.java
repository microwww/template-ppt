package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.replace.ReplaceExpress;
import com.github.microwww.ttp.replace.SearchTable;
import com.github.microwww.ttp.replace.SearchTableCell;
import com.github.microwww.ttp.replace.SearchTableRow;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class ReplaceOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(ReplaceOperation.class);

    public ReplaceOperation() {
    }

    @Override
    public void parse(ParseContext context) {
        context.getTemplate();
        List<?> search = super.search(context);
        for (Object o : search) {
            thisInvoke("replace", new Object[]{context, o});
        }
    }

    // DEFAULT
    public void replace(ParseContext context, Object item) {
        logger.warn("Not support type : {}", item.getClass());
    }

    public void replace(ParseContext context, XSLFTable item) {
        List<ReplaceExpress> list = new SearchTable(item).search();
        replace(context, list);
    }

    public void replace(ParseContext context, XSLFTableRow item) {
        List<ReplaceExpress> list = new SearchTableRow(item).search();
        replace(context, list);
    }

    public void replace(ParseContext context, XSLFTableCell item) {
        if (this.getParams().length > 0) {
            String param = this.getParams()[0];
            String val = getValue(param, context.getData()).toString();
            try {
                item.getTextParagraphs().get(0).getTextRuns().get(0).setText(val);
            } catch (Exception ex) {
                item.addNewTextParagraph().addNewTextRun();
                item.getTextParagraphs().get(0).getTextRuns().get(0).setText(val);
            }
        } else {
            List<ReplaceExpress> list = new SearchTableCell(item).search();
            replace(context, list);
        }
    }

    private void replace(ParseContext context, List<ReplaceExpress> list) {
        for (ReplaceExpress express : list) {
            String exp = express.getExpress();
            String val = getValue(exp, context.getData()).toString();
            express.replace(val);
        }
    }

}
