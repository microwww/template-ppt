package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public interface SearchContent {

    public static final Pattern PATTERN = Pattern.compile("\\$\\{[$!=<~>a-zA-Z._0-9/\\[\\]@*?\\(\\)]+\\}");

    List<TextExpress> search();

    public static List<TextExpress> searchExpress(XSLFTextRun run) {
        String text = run.getRawText();
        Matcher matcher = PATTERN.matcher(text);
        List<TextExpress> list = new ArrayList<>();
        while (matcher.find()) {
            String group = matcher.group();
            list.add(new TextExpress(run, group, group.substring(2, group.length() - 1)));
        }
        return list;
    }
}
