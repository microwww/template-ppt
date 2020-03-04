package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.xslf.XSLFGraphicChart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Stack;
import java.util.concurrent.ConcurrentHashMap;

public class ParseContext {
    private static final Logger logger = LoggerFactory.getLogger(ParseContext.class);
    private final Map<String, Class> supportShape = new ConcurrentHashMap<>();

    {
        Class<XSLFGraphicChart> clazz = XSLFGraphicChart.class;
        supportShape.put(clazz.getSimpleName(), clazz);
    }

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

    public ParseContext parse(InputStream format, Class<Operation>... support) throws IOException {
        return this.parse(new InputStreamReader(format, "UTF-8"), support);
    }

    public ParseContext parse(InputStreamReader format, Class<Operation>... support) throws IOException {
        ParseExpresses exp = new ParseExpresses().addSupportOperations(support);
        List<Operation> opts = exp.parse(format);

        for (Operation opt : opts) {
            try {
                opt.parse(this);
            } catch (RuntimeException e) {
                logger.error("OPERATION : {} ( {} )", opt.getPrefix(), opt.getNode(), opt.getParams());
                throw e;
            }
        }
        return this;
    }

    public void write(OutputStream out) throws IOException {
        this.getTemplateShow().write(out);
    }

    public void putSupportShape(String name, Class clazz) {
        supportShape.put(name, clazz);
    }

    public Class parseShapeClass(String exp) {
        Class clazz = supportShape.get(exp);
        if (clazz == null) {
            String cname = "org.apache.poi.xslf.usermodel." + exp;
            try {
                return Class.forName(cname);
            } catch (ClassNotFoundException e) {// IGNORE
            }
        }
        throw new UnsupportedOperationException("Not support type : " + exp);
    }
}
