package com.github.microwww.ttp;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.IOException;
import java.util.List;

public class UtilPPT {

    private static final Logger log = LoggerFactory.getLogger(UtilPPT.class);

    public static XSLFTable copeTable(XSLFSlide slide, XSLFTable src) {
        return _Help.copeTable(slide, src);
    }

    public static XSLFTableRow copeTableRow(XSLFTable table, XSLFTableRow src) {
        return _Help.copyTableRow(table, src);
    }

    /**
     * @param slide
     * @param index from 0
     * @return maybe null
     */
    public static XSLFChart findChart(XSLFSlide slide, int index) {
        int i = 0;
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part instanceof XSLFChart) {
                if (i == index) {
                    return (XSLFChart) part;
                }
                i++;
            }
        }
        return null; //throw new IllegalStateException("chart not found in the template");
    }

    private static int px2point(double px) {
        return (int) (Math.rint(px * Units.EMU_PER_POINT));
    }

    public static Rectangle rectanglePx2point(Rectangle2D px) {
        return new Rectangle(px2point(px.getX()), px2point(px.getY()), px2point(px.getWidth()), px2point(px.getHeight()));
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
        values1[6] = 16.0; // if you ever want to change the underlying data
        final XDDFNumericalDataSource<? extends Number> valuesData2 = XDDFDataSourcesFactory.fromArray(values2, valuesDataRange2, 2);

        XDDFChartData.Series series1 = bar.getSeries().get(0);
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

    public static File createCanWriteDirection(File file) throws IOException {
        if (!file.exists()) {
            if (!file.mkdirs()) {
                log.warn("Make dir error ! {} ", file.getCanonicalPath());
            }
        }
        if (!file.canWrite()) {
            throw new IOException("Can not to write file to : " + file.getCanonicalFile());
        }
        file.mkdirs();
        return file;
    }
}
