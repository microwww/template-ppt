package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.Tools;
import com.github.microwww.ttp.replace.ReplaceExpress;
import com.github.microwww.ttp.replace.SearchTable;
import com.github.microwww.ttp.replace.SearchTableCell;
import com.github.microwww.ttp.replace.SearchTableRow;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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

    public void replace(ParseContext context, XSLFTable item) {
        List<ReplaceExpress> list = new SearchTable(item).search();
        replace(context, list);
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

        String title = super.getValue(params[0], context.getData()).toString();

        List<Object> categories = (List) super.getValue(params[1], context.getData());
        String[] cts = parse2string(categories);

        if (type instanceof XDDFPieChartData) {
            List<Object> values = (List) super.getValue(params[2], context.getData());
            Assert.isTrue(values.size() == categories.size(), "Error CATEGORY.length != VALUE.length");
            Double[] dbs = parse2double(values);
            Tools.setPieDate(chart, title, cts, dbs);
        } else if (type instanceof XDDFRadarChartData) {
            List<Object> series = (List) super.getValue(params[2], context.getData());
            String[] ss = parse2string(series);
            Double[][] dbs = new Double[params.length - 3][];
            for (int i = 0; i < dbs.length; i++) {
                List<Object> values = (List) super.getValue(params[i + 3], context.getData());
                dbs[i] = parse2double(values);
            }
            Tools.setRadarData(chart, title, cts, ss, dbs);
        }

    }

    public static Double[] parse2double(List<Object> values) {
        Double[] dbs = new Double[values.size()];
        for (int i = 0; i < dbs.length; i++) {
            dbs[i] = Double.valueOf(values.get(i).toString());
        }
        return dbs;
    }

    public static String[] parse2string(List<Object> values) {
        String[] cts = new String[values.size()];
        for (int i = 0; i < cts.length; i++) {
            cts[i] = values.get(i).toString();
        }
        return cts;
    }

    public void replace(ParseContext context, XSLFTableRow item) {
        List<ReplaceExpress> list = new SearchTableRow(item).search();
        replace(context, list);
    }

    public void replace(ParseContext context, XSLFTableCell item) {
        if (this.getParams().length > 0) {
            String param = this.getParams()[0];
            String val = getValue(param, context.getData()).toString();
            Tools.setCellTextWithStyle(item, val);
        } else {
            List<ReplaceExpress> list = new SearchTableCell(item).search();
            replace(context, list);
        }
    }

    private void replace(ParseContext context, List<ReplaceExpress> list) {
        for (ReplaceExpress express : list) {
            String exp = express.getExpress();
            String val = getValue(exp, context.getData()).toString();
            express.replace(val);
        }
    }

}
