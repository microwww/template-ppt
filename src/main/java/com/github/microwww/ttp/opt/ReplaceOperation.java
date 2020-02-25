package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.Tools;
import com.github.microwww.ttp.replace.ReplaceExpress;
import com.github.microwww.ttp.replace.SearchContent;
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
            thisInvoke("replace", new Object[]{context, o});
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
        Object value = super.getValue(msg.getParam(), context.getData());
        String title = msg.format(value);

        Collection categories = super.getCollectionValue(params[1], context.getData());
        String[] cts = parse2string(categories);

        if (type instanceof XDDFPieChartData) {
            Collection values = super.getCollectionValue(params[2], context.getData());
            Assert.isTrue(values.size() == categories.size(), "Error CATEGORY.length != VALUE.length");
            Double[] dbs = parse2double(values);
            Tools.setPieDate(chart, title, cts, dbs);
        } else if (type instanceof XDDFRadarChartData) {
            Collection series = super.getCollectionValue(params[2], context.getData());
            String[] ss = parse2string(series);
            Double[][] dbs = new Double[params.length - 3][];
            for (int i = 0; i < dbs.length; i++) {
                Collection values = super.getCollectionValue(params[i + 3], context.getData());
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

    public void replace(ParseContext context, XSLFTableRow item) {
        List<ReplaceExpress> list = SearchContent.search(item);
        for (XSLFTableCell cell : item.getCells()) {
            replace(context, cell);
        }
    }

    public void replace(ParseContext context, XSLFTextShape item) {
        if (this.getParams().length == 1) {
            StringBuffer buffer = new StringBuffer();
            for (ParamMessage param : this.getParamsWithPattern()) {
                Object val = getValue(param.getParam(), context.getData());
                buffer.append(param.format(val));
            }
            Tools.setTextShapeWithStyle(item, buffer.toString());
        } else {
            List<ReplaceExpress> search = SearchContent.search(item);
            Object[] vals = new Object[this.getParams().length];
            for (int i = 0; i < vals.length; i++) {
                vals[i] = this.getValue(this.getParams()[i], context.getData());
            }

            for (ReplaceExpress express : search) {
                String pattern = express.getPattern();
                String val = new ParamMessage(null, pattern).format(vals);
                express.replace(val);
            }

        }
    }

}
