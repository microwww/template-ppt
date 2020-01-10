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
            Rectangle2D point = rectanglePx2point(val.getKey().getAnchor(), delta.getX(), delta.getY(), delta.getWidth(), delta.getHeight());
            dest.addChart(chart, point);
        });
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, XSLFSheet dest, Rectangle position) {
        return _Help.copyChart(src, index, dest, (chart, val) -> {
            dest.addChart(chart, position); //rectanglePx2point(val.getKey().getAnchor(), 0, 0, 0, 0));
        });
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

    public static void setBarData(XSLFChart chart, String chartTitle, String[] series, String[] categories, Double[] values1, Double[] values2) {
        final List<XDDFChartData> data = chart.getChartSeries();
        final XDDFBarChartData bar = (XDDFBarChartData) data.get(0);

        final int numOfPoints = categories.length;
        final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        final String valuesDataRange2 = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
        final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values1, valuesDataRange, 1);
        final XDDFNumericalDataSource<? extends Number> valuesData2 = XDDFDataSourcesFactory.fromArray(values2, valuesDataRange2, 2);

        XDDFChartData.Series series1 = bar.getSeries(0);
        series1.replaceData(categoriesData, valuesData);
        series1.setTitle(series[0], chart.setSheetTitle(series[0], 0));
        XDDFChartData.Series series2 = bar.addSeries(categoriesData, valuesData2);
        series2.setTitle(series[1], chart.setSheetTitle(series[1], 1));

        chart.plot(bar);
        chart.setTitleText(chartTitle); // https://stackoverflow.com/questions/30532612
        // chart.setTitleOverlay(overlay);
    }

    public static void setPieDate(XSLFChart chart, String chartTitle, String[] categories, Double[] values) {
        // Series Text
        List<XDDFChartData> series = chart.getChartSeries();
        XDDFPieChartData pie = (XDDFPieChartData) series.get(0);

        final int numOfPoints = categories.length;
        final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange);
        final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values, valuesDataRange);

        XDDFPieChartData.Series firstSeries = (XDDFPieChartData.Series) pie.getSeries(0);
        firstSeries.replaceData(categoriesData, valuesData);
        firstSeries.setTitle(chartTitle, chart.setSheetTitle(chartTitle, 0));
        // firstSeries.setExplosion(25);
        chart.plot(pie);
    }
}
