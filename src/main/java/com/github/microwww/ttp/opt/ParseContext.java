package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSheet;

import java.util.HashMap;
import java.util.Map;
import java.util.Stack;

public class ParseContext {

    private Stack<Object> container = new Stack<>();
    private Object data = new HashMap<>();

    public ParseContext(XMLSlideShow template) {
        container.add(template);
    }

    public XMLSlideShow getTemplateShow() {
        return (XMLSlideShow) container.get(0);
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

    public void setTemplate(XSLFSheet sheet) {
        XMLSlideShow show = this.getTemplateShow();
        container.clear();
        container.push(show);
        container.push(sheet);
    }

    public XSLFSheet getTemplate() {
        for (Object o : container) {
            if (o instanceof XSLFSheet) {
                return (XSLFSheet) o;
            }
        }
        throw new RuntimeException("Not find XSLFSheet !");
    }

    public Stack<Object> getContainer() {
        return container;
    }
}
