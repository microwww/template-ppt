package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XSLFSheet;

public class RepeatOperation extends Operation {

    @Override
    public void parse(ParseContext context) {
        XSLFSheet slide = context.getTemplate();
        super.search(context);
    }
}
