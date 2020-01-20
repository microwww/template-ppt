package org.apache.poi.xslf.usermodel;

import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;

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
