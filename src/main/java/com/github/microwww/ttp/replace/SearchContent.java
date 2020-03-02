package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.Format;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;

public abstract class SearchContent {

    //public static final Pattern PATTERN = Pattern.compile("\\$\\{[0-9]+(,[a-zA-Z]+(,.*+)?)?\\}");
    public static final Logger logger = LoggerFactory.getLogger(SearchContent.class);

    public static List<ReplaceExpress> searchExpress(XSLFTextParagraph run) {
        String pattern = run.getText();
        List<ReplaceExpress> list = new ArrayList<>();
        for (int i = 0; i < pattern.length(); ) {
            int idx = pattern.indexOf("${", i);
            if (idx < 0) {
                break;
            }
            int next = pattern.indexOf('}', idx);
            if (next < 0) {
                break;
            }
            String search = pattern.substring(idx, next + 1);
            list.add(new ReplaceExpress(run, search));
            i = next;
        }
        Format[] formats = new MessageFormat(pattern).getFormats();
        if (list.size() != formats.length) {
            logger.warn("Please check your MessageFormat text !");
        }
        return list;
    }

    public static List<ReplaceExpress> search(XSLFTextShape item) {
        List<ReplaceExpress> list = new ArrayList<>();
        List<XSLFTextParagraph> pgs = item.getTextParagraphs();
        for (XSLFTextParagraph pg : pgs) {
            list.addAll(SearchContent.searchExpress(pg));
        }
        return list;
    }

    public static List<ReplaceExpress> search(XSLFTableRow row) {
        ArrayList<ReplaceExpress> list = new ArrayList<>();
        for (XSLFTableCell cs : row.getCells()) {
            List<ReplaceExpress> search = SearchContent.search(cs);
            list.addAll(search);
        }
        return list;
    }

    public static List<ReplaceExpress> search(XSLFTable table) {
        ArrayList<ReplaceExpress> list = new ArrayList<>();
        for (XSLFTableRow row : table.getRows()) {
            List<ReplaceExpress> search = SearchContent.search(row);
            list.addAll(search);
        }
        return list;
    }
}
