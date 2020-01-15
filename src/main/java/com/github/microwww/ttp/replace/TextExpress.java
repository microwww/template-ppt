package com.github.microwww.ttp.replace;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

public class TextExpress {
    private XSLFTextRun run;
    private String text;
    private String express;

    public TextExpress() {
    }

    public TextExpress(XSLFTextRun run, String text, String express) {
        this.run = run;
        this.text = text;
        this.express = express;
    }

    public void replace(String text) {
        run.setText(StringUtils.replace(run.getRawText(), this.text, text));
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public String getExpress() {
        return express;
    }

    public void setExpress(String express) {
        this.express = express;
    }
}