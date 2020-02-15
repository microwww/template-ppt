package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public interface SearchContent {

    public static final Pattern PATTERN = Pattern.compile("\\$\\{[$!=<~>a-zA-Z._0-9/\\[\\]@*?\\(\\)]+\\}");

    List<ReplaceExpress> search();

    public static List<ReplaceExpress> searchExpress(XSLFTextRun run) {
        String text = run.getRawText();
        Matcher matcher = PATTERN.matcher(text);
        List<ReplaceExpress> list = new ArrayList<>();
        while (matcher.find()) {
            String group = matcher.group();
            list.add(new ReplaceExpress(run, group, group.substring(2, group.length() - 1)));
        }
        return list;
    }
}
