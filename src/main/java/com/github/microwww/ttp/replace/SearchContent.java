package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public interface SearchContent {

    public static final Pattern PATTERN = Pattern.compile("\\$\\{[$!=<~>a-zA-Z._0-9/\\[\\]@*?\\(\\)]+\\}");

    List<ReplaceExpress> search();

    public static List<ReplaceExpress> searchExpress(XSLFTextParagraph run) {
        String text = run.getText();
        Matcher matcher = PATTERN.matcher(text);
        List<ReplaceExpress> list = new ArrayList<>();
        while (matcher.find()) {
            String group = matcher.group();
            list.add(new ReplaceExpress(run, group, group.substring(2, group.length() - 1)));
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
