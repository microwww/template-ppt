package com.github.microwww.ttp.operation;

import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFTable;

public class Chart implements ElementsOperation<XSLFChart, XSLFSheet> {

    @Override
    public void copy(XSLFChart src, XSLFSheet shapes) {
    }

    @Override
    public void delete(XSLFChart src, XSLFSheet shapes) {
    }

    @Override
    public void replace(XSLFChart src, XSLFSheet shapes) {
    }
}
