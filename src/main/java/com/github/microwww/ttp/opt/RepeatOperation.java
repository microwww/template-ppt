package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XSLFSheet;

public class RepeatOperation extends Operation {

    public boolean isSupport() {
        return getExpresses()[0].equalsIgnoreCase("repeat");
    }

    public void parse(XSLFSheet slide) throws ClassNotFoundException {
        super.searchElement(slide, 1);
    }
}
