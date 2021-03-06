package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.Tools;
import com.github.microwww.ttp.replace.ReplaceExpress;
import com.github.microwww.ttp.replace.SearchContent;
import com.github.microwww.ttp.xslf.XSLFGraphicChart;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Collection;
import java.util.List;

public class ReplaceOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(ReplaceOperation.class);

    public ReplaceOperation() {
    }

    @Override
    public void parse(ParseContext context) {
        List<?> search = super.search(context);
        for (Object o : search) {
            thisInvoke("replace", context, o);
        }
    }

    // DEFAULT
    public void replace(ParseContext context, Object item) {
        logger.warn("Not support type : {}", item.getClass());
    }

    public void replace(ParseContext context, XSLFGraphicChart item) {
        XSLFChart chart = item.getChart();
        List<XDDFChartData> data = chart.getChartSeries();
        if (data.isEmpty()) {
            throw new IllegalArgumentException("NO know chart type");
        }
        XDDFChartData type = data.get(0);
        String[] params = this.getParams();
        Assert.isTrue(params.length >= 3, "Chart data [title, category[], data[]], Must 3 params value");

        ParamMessage msg = this.getParamsWithPattern()[0];
        Object value = super.getValue(msg.getParam(), context.getDataStack());
        String title = msg.format(value);

        List categories = super.getCollectionValue(params[1], context.getDataStack());
        String[] cts = parse2string(categories);

        if (type instanceof XDDFPieChartData) {
            List values = super.getCollectionValue(params[2], context.getDataStack());
            Assert.isTrue(values.size() == categories.size(), "Error CATEGORY.length != VALUE.length");
            Double[] dbs = parse2double(values);
            Tools.setPieDate(chart, title, cts, dbs);
        } else {
            if (!(type instanceof XDDFRadarChartData) && !(type instanceof XDDFBarChartData)){
                logger.warn("UNKNOWN Chart type : {}, ", type);
            }
            List series = super.getCollectionValue(params[2], context.getDataStack());
            String[] ss = parse2string(series);
            Double[][] dbs = new Double[params.length - 3][];
            for (int i = 0; i < dbs.length; i++) {
                List values = super.getCollectionValue(params[i + 3], context.getDataStack());
                dbs[i] = parse2double(values);
            }
            Tools.setRadarData(chart, title, cts, ss, dbs);
        }

    }

    public static Double[] parse2double(Collection<Object> values) {
        Double[] dbs = new Double[values.size()];
        int i = 0;
        for (Object value : values) {
            dbs[i++] = Double.valueOf(value.toString());
        }
        return dbs;
    }

    public static String[] parse2string(Collection<Object> values) {
        String[] cts = new String[values.size()];
        int i = 0;
        for (Object value : values) {
            cts[i++] = value.toString();
        }
        return cts;
    }

    public void replace(ParseContext context, XSLFTextParagraph paragraph) {
        List<ReplaceExpress> exps = SearchContent.searchExpress(paragraph);
        if (exps.isEmpty()) {
            StringBuilder buffer = new StringBuilder();
            for (ParamMessage param : this.getParamsWithPattern()) {
                Object val = getValue(param.getParam(), context.getDataStack());
                buffer.append(param.format(val));
            }
            Tools.setParagraphText(paragraph, buffer.toString());
        } else {
            this.writeShape(context, exps);
        }
    }

    public void replace(ParseContext context, XSLFTableRow item) {
        for (XSLFTableCell cell : item.getCells()) {
            replace(context, cell);
        }
    }

    public void replace(ParseContext context, XSLFTextBox box) {
        this.replace(context, (XSLFTextShape) box);
    }

    public void replace(ParseContext context, XSLFTableCell cell) {
        this.replace(context, (XSLFTextShape) cell);
    }

    public void replace(ParseContext context, XSLFTextShape item) {
        if (this.getParams().length == 1) {
            StringBuilder buffer = new StringBuilder();
            for (ParamMessage param : this.getParamsWithPattern()) {
                Object val = getValue(param.getParam(), context.getDataStack());
                buffer.append(param.format(val));
            }
            Tools.setTextShapeWithStyle(item, buffer.toString());
        } else {
            List<ReplaceExpress> search = SearchContent.search(item);
            writeShape(context, search);
        }
    }

    public void writeShape(ParseContext context, List<ReplaceExpress> search) {
        Object[] vals = new Object[this.getParams().length];
        for (int i = 0; i < vals.length; i++) {
            vals[i] = this.getValue(this.getParams()[i], context.getDataStack());
        }

        for (ReplaceExpress express : search) {
            String pattern = express.getPattern();
            String val = new ParamMessage(null, pattern).format(vals);
            express.replace(val);
        }
    }

}
