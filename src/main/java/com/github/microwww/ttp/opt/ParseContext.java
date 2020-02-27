package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSheet;

import java.util.HashMap;
import java.util.Map;
import java.util.Stack;

public class ParseContext {

    private final Stack<Object> container = new Stack<>();
    private final Stack<Object> data = new Stack<>();

    public ParseContext(XMLSlideShow template) {
        container.add(template);
        data.push(new HashMap<>());
    }

    public XMLSlideShow getTemplateShow() {
        return (XMLSlideShow) container.get(0);
    }

    public Object getData() {
        return data.peek();
    }

    public void setData(Object data) {
        this.data.clear();
        this.data.push(data);
    }

    public void addData(String value, Object data) {
        Object peek = this.data.peek();
        if (peek instanceof Map) {
            ((Map<String, Object>) peek).put(value, data);
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
