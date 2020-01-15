package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;

import java.util.ArrayList;
import java.util.List;

public class SearchTableRow implements SearchContent {
    private final XSLFTableRow row;

    public SearchTableRow(XSLFTableRow row) {
        this.row = row;
    }

    @Override
    public List<TextExpress> search() {
        ArrayList<TextExpress> list = new ArrayList<>();
        for (XSLFTableCell cs : row.getCells()) {
            List<TextExpress> search = new SearchTableCell(cs).search();
            list.addAll(search);
        }
        return list;
    }
}