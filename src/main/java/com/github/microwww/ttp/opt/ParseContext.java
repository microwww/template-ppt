package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Stack;

public class ParseContext {
    private static final Logger logger = LoggerFactory.getLogger(ParseContext.class);

    private final Stack<Object> container = new Stack<>();
    private final Stack<Object> data = new Stack<>();

    public ParseContext(XMLSlideShow template) {
        container.add(template);
        if (!template.getSlides().isEmpty()) {
            container.push(template.getSlides().get(0));
        }
        data.push(new HashMap<>());
    }

    public XMLSlideShow getTemplateShow() {
        return (XMLSlideShow) container.get(0);
    }

    public Stack<Object> getDataStack() {
        return data;
    }

    public Stack<Object> pushData(Object data) {
        this.data.push(data);
        return this.data;
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
        for (int i = container.size(); i > 0; i--) {
            Object o = container.get(i - 1);
            if (o instanceof XSLFSheet) {
                return (XSLFSheet) o;
            }
        }
        throw new RuntimeException("Not find XSLFSheet !");
    }

    public Stack<Object> getContainer() {
        return container;
    }

    public void parse(InputStream format, OutputStream out) throws IOException {
        ParseExpresses exp = new ParseExpresses(format);
        exp.parse();
        List<Operation> opts = exp.getOperations();

        for (Operation opt : opts) {
            try {
                opt.parse(this);
            } catch (RuntimeException e) {
                logger.error("OPERATION : {} ( {} )", opt.getPrefix(), opt.getNode(), opt.getParams());
                throw e;
            }
        }

        this.getTemplateShow().write(out);
    }
}
