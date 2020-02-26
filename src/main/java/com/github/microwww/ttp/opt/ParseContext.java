package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSheet;

import java.util.HashMap;
import java.util.Map;

public class ParseContext {

    private static final String name_prefix = ParseContext.class.getName();
    public static final String TEMPLATE_SHOW = name_prefix + ".TEMPLATE_SHOW";
    public static final String TEMPLATE = name_prefix + ".TEMPLATE";
    //public static final String TARGET_SHOW = name_prefix + ".TARGET_SHOW";
    //public static final String TARGET = name_prefix + ".TARGET";

    private Object data = new HashMap<>();

    public ParseContext(XMLSlideShow template) {
        map.put(TEMPLATE_SHOW, template);
    }

    private Map<String, Object> map = new HashMap<>();

    public XMLSlideShow getTemplateShow() {
        return (XMLSlideShow) map.get(TEMPLATE_SHOW);
    }

    public void setTemplateShow(XMLSlideShow templateShow) {
        map.put(TEMPLATE_SHOW, templateShow);
    }

    public XSLFSheet getTemplate() {
        return (XSLFSheet) map.get(TEMPLATE);
    }

    public void setTemplate(XSLFSheet template) {
        map.put(TEMPLATE, template);
    }

    public void putConifg(String key, Object val) {
        map.put(key, val);
    }

    public <T> T getConfig(String key, Class<T> t) {
        return (T) map.get(key);
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public void addData(String value, Object data) {
        if (this.data instanceof Map) {
            ((Map) this.data).put(value, data);
        } else {
            throw new UnsupportedOperationException();
        }
    }
}
