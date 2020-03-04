package com.github.microwww.ttp.xslf;

import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;

public class XSLFGraphicChart {
    private final XSLFGraphicFrame graphic;
    private final XSLFChart chart;

    public XSLFGraphicChart(XSLFGraphicFrame graphic, XSLFChart chart) {
        this.graphic = graphic;
        this.chart = chart;
    }

    public XSLFGraphicFrame getGraphic() {
        return graphic;
    }

    public XSLFChart getChart() {
        return chart;
    }
}
