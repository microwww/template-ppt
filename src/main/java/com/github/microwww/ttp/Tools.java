package com.github.microwww.ttp;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.util.List;

public class Tools {

    private static final Logger log = LoggerFactory.getLogger(Tools.class);

    public static XSLFTable copyTable(XSLFSheet sheet, XSLFTable src) {
        return _Help.copyTable(sheet, src);
    }

    public static XSLFTableRow copyTableRow(XSLFTable table, XSLFTableRow src) {
        return _Help.copyTableRow(table, src);
    }

    public static XSLFAutoShape copyAutoShape(XSLFSheet sheet, XSLFAutoShape src) {
        XSLFAutoShape shape = sheet.createAutoShape();
        shape.getXmlObject().set(src.getXmlObject().copy());
        //Rectangle2D position = src.getAnchor();
        //Rectangle2D.Double rectangle = new Rectangle2D.Double(position.getX(), position.getY(), position.getWidth(), position.getHeight());
        //shape.setAnchor(rectangle);
        return shape;
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, XSLFSheet dest) {
        return _Help.copyChart(src, index, dest);
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, XSLFSheet dest, Rectangle2D delta) {
        return _Help.copyChart(src, index, dest, (chart, val) -> {
            Rectangle2D point = rectanglePx2point(val.getGraphic().getAnchor(), delta.getX(), delta.getY(), delta.getWidth(),
                    delta.getHeight());
            dest.addChart(chart, point);
        });
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, XSLFSheet dest, Rectangle position) {
        return _Help.copyChart(src, index, dest, (chart, val) -> {
            dest.addChart(chart, position); //rectanglePx2point(val.getKey().getAnchor(), 0, 0, 0, 0));
        });
    }

    public static boolean isEqual(double a, double b, double delta) {
        return Math.abs(a - b) < Math.abs(delta);
    }

    /**
     * @param sheet target slide
     * @param index from 0
     * @return maybe null
     */
    public static XSLFChart findChart(XSLFSheet sheet, int index) {
        return _Help.findChart(sheet, index);
    }

    public static Rectangle rectanglePx2point(Rectangle2D px) {
        return _Help.rectanglePx2point(px, px.getX(), px.getY(), px.getWidth(), px.getHeight());
    }

    public static Rectangle2D delta(Rectangle2D px, double x, double y, double w, double h) {
        return new Rectangle2D.Double(px.getX() + x, px.getY() + y, px.getWidth() + w, px.getHeight() + h);
    }

    public static Rectangle rectanglePx2point(Rectangle2D px, double x, double y, double w, double h) {
        return _Help.rectanglePx2point(px, x, y, w, h);
    }

    public static void setTextShapeWithStyle(XSLFTextShape item, String val) {
        for (int i = item.getTextParagraphs().size() - 1; i > 0; i--) {
            item.getTextBody().removeParagraph(i);
        }
        if (item.getTextParagraphs().isEmpty()) {
            item.addNewTextParagraph().addNewTextRun();
        }
        XSLFTextParagraph paragraph = item.getTextParagraphs().get(0);
        List<XSLFTextRun> runs = paragraph.getTextRuns();

        for (int i = 1; i < runs.size(); i++) {
            runs.get(i).setText("");
        }
        if (runs.isEmpty()) {
            paragraph.addNewTextRun();
        }
        runs.get(0).setText(val);
    }

    public static void setRadarData(XSLFChart chart, String chartTitle, String[] series, String[] categories, Double[]... values) {
        int size = categories.length;
        List<XDDFChartData> s = chart.getChartSeries();
        XDDFChartData bar = s.get(0);

        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, size, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(//
                categories, categoryDataRange, 0);

        Assert.isTrue(series.length == values.length, "Error : series.length != values.length");
        for (int i = 0; i < series.length; i++) {
            String valuesDataRange = chart.formatRange(new CellRangeAddress(1, size, i + 1, i + 1));
            XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values[i], valuesDataRange, 1);
            List<XDDFChartData.Series> seriess = bar.getSeries();
            XDDFChartData.Series ss;
            if (seriess.size() > i) {
                ss = seriess.get(i);
                ss.replaceData(categoriesData, valuesData);
            } else {
                ss = bar.addSeries(categoriesData, valuesData);
            }
            ss.setTitle(series[i], chart.setSheetTitle(series[i], 1));
        }

        chart.plot(bar);
        chart.setTitleText(chartTitle); // https://stackoverflow.com/questions/30532612
    }

    public static void setPieDate(XSLFChart chart, String chartTitle, String[] categories, Double[] values) {
        // Series Text
        List<XDDFChartData> series = chart.getChartSeries();
        XDDFPieChartData pie = (XDDFPieChartData) series.get(0);

        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange);
        XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values, valuesDataRange);

        XDDFPieChartData.Series firstSeries = (XDDFPieChartData.Series) pie.getSeries(0);
        firstSeries.replaceData(categoriesData, valuesData);
        firstSeries.setTitle(chartTitle, chart.setSheetTitle(chartTitle, 0));
        // firstSeries.setExplosion(25);
        chart.plot(pie);
    }
}
