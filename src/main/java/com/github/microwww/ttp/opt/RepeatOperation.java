package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XSLFSheet;

import java.util.List;

public class RepeatOperation extends Operation {

    @Override
    public void parse(XSLFSheet slide, List<Operation> parsed) {
        try {
            super.searchElement(slide, 1);
        } catch (ClassNotFoundException e) {
            throw new UnsupportedOperationException("Un supoort type :: " + e.getMessage(), e);
        }
    }
}
