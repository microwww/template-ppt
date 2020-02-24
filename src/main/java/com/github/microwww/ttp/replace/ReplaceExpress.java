package com.github.microwww.ttp.replace;

import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;

public class ReplaceExpress {
    private XSLFTextParagraph run;
    private String text;

    public ReplaceExpress(XSLFTextParagraph run, String text) {
        this.run = run;
        this.text = text;
    }

    public String getPattern() {
        return text.substring(1);
    }

    public void replace(String text) {
        Tools.replace(run, this.text, text);
        //run.setText(StringUtils.replace(run.getRawText(), this.text, text));
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

}