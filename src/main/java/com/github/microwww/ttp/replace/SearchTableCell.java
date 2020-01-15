package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.util.Collections;
import java.util.List;

public class SearchTableCell implements SearchContent {
    private final XSLFTableCell cell;

    public SearchTableCell(XSLFTableCell cell) {
        this.cell = cell;
    }

    @Override
    public List<TextExpress> search() {
        List<XSLFTextParagraph> pgs = this.cell.getTextParagraphs();
        for (XSLFTextParagraph pg : pgs) {
            for (XSLFTextRun run : pg.getTextRuns()) {
                return SearchContent.searchExpress(run);
            }
        }
        return Collections.emptyList();
    }

}