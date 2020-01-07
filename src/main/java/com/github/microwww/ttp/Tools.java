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

    public static XSLFTable copyTable(XSLFSlide slide, XSLFTable src) {
        return _Help.copyTable(slide, src);
    }

    public static XSLFTableRow copyTableRow(XSLFTable table, XSLFTableRow src) {
        return _Help.copyTableRow(table, src);
    }

    public static XSLFChart copyChart(XSLFSlide src, int index, XSLFSlide dest) {
        return _Help.copyChart(src, index, dest);
    }

    /**
     * @param slide
     * @param index from 0
     * @return maybe null
     */
    public static XSLFChart findChart(XSLFSlide slide, int index) {
        return _Help.findChart(slide, index);
    }

    public static Rectangle rectanglePx2point(Rectangle2D px) {
        return _Help.rectanglePx2point(px);
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

        XDDFPieChartData.Series firstSeries = (XDDFPieChartData.Series) pie.getSeries().get(0);
        firstSeries.replaceData(categoriesData, valuesData);
        firstSeries.setTitle(chartTitle, chart.setSheetTitle(chartTitle, 0));
        // firstSeries.setExplosion(25);
        chart.plot(pie);
    }
}
