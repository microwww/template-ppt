package com.github.microwww.ttp;

import com.github.microwww.ttp.util.BiConsumer;
import com.github.microwww.ttp.util._Help;
import com.github.microwww.ttp.xslf.XSLFGraphicChart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class Tools {

    private static final Logger log = LoggerFactory.getLogger(Tools.class);

    public static XSLFTable copyTable(XSLFSheet sheet, XSLFTable src) {
        return _Help.copyTable(sheet, src);
    }

    public static XSLFTableRow copyTableRow(XSLFTable table, XSLFTableRow src) {
        return _Help.copyTableRow(table, src);
    }

    public static XSLFTable getTable(XSLFTableRow src) {
        return _Help.getTable(src);
    }

    public static XSLFTable getTable(XSLFTableCell src) {
        return _Help.getTable(src);
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

    public static XSLFChart copyChart(XSLFSheet src, int index, final XSLFSheet dest, final Rectangle2D delta) {
        return _Help.copyChart(src, index, dest, new BiConsumer<XSLFChart, XSLFGraphicChart>() {
            public void accept(XSLFChart chart, XSLFGraphicChart val) {
                Rectangle2D point = rectanglePx2point(val.getGraphic().getAnchor(), delta.getX(), delta.getY(), delta.getWidth(),
                        delta.getHeight());
                dest.addChart(chart, point);
            }
        });
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, final XSLFSheet dest, final Rectangle position) {
        return _Help.copyChart(src, index, dest, new BiConsumer<XSLFChart, XSLFGraphicChart>() {
            public void accept(XSLFChart chart, XSLFGraphicChart val) {
                dest.addChart(chart, position); //rectanglePx2point(val.getKey().getAnchor(), 0, 0, 0, 0));
            }
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
        return rectanglePx2point(px, px.getX(), px.getY(), px.getWidth(), px.getHeight());
    }

    public static Rectangle2D delta(Rectangle2D px, double x, double y, double w, double h) {
        return new Rectangle2D.Double(px.getX() + x, px.getY() + y, px.getWidth() + w, px.getHeight() + h);
    }

    public static Rectangle rectanglePx2point(Rectangle2D px, double x, double y, double w, double h) {
        return rectanglePx2point(px, x, y, w, h);
    }

    public static void setTextShapeWithStyle(XSLFTextShape item, String val) {
        List<XSLFTextParagraph> ghs = item.getTextParagraphs();
        if (ghs.isEmpty()) {
            throw new IllegalArgumentException("You must write some word for Get text format");
        }
        for (int i = ghs.size() - 1; i > 0; i--) {
            item.getTextBody().removeParagraph(i);
        }
        XSLFTextParagraph paragraph = ghs.get(0);
        setParagraphText(paragraph, val);
    }

    public static void setParagraphText(XSLFTextParagraph paragraph, String val) {
        replace(paragraph, paragraph.getText(), val);
    }

    public static void replace(XSLFTextParagraph paragraph, String origin, String replacement) {
        CTTextParagraph xml = paragraph.getXmlObject();
        StringBuffer buffer = new StringBuffer();
        List<CTRegularTextRun> runs = xml.getRList();
        if (runs.isEmpty()) {
            log.debug("Not find run . set some text is good !");
            xml.addNewR().setT("-");
            runs = paragraph.getXmlObject().getRList();
        }
        for (CTRegularTextRun rText : runs) {
            buffer.append(rText.getT());
        }
        int idx = buffer.indexOf(origin), fromRun = -1, toRun = -1, cursor = 0;
        int end = idx + origin.length();
        for (int i = 0; i < runs.size(); i++) {
            CTRegularTextRun rText = runs.get(i);
            String text = rText.getT(), start = "", fix = "";
            int len = text.length();

            if (idx >= cursor + len) {// 未开始
                cursor += len;
                continue;
            }
            if (end < cursor) { // 已结束
                break;
            }

            if (idx >= cursor && idx <= cursor + len) {// 起始
                fromRun = i;
                // int fromRunPosition = idx - cursor;
                start = text.substring(0, idx - cursor) + replacement;// 开始替换:取前半段然后追加替换字符
            }
            if (end >= cursor && end <= cursor + len) {
                toRun = i;
                // int toRunPosition = end - cursor;
                fix = text.substring(end - cursor); // 取后半段
            }
            rText.setT(start + fix);
            cursor += len;
        }
        if (fromRun >= 0 && toRun >= 0) {
            log.debug("NOT find message by XML : {}, check it", origin);
        }
    }

    public static XSLFTextParagraph copy(XSLFTextParagraph paragraph) {
        XSLFTextShape shape = paragraph.getParentShape();
        XSLFTextParagraph newPg = shape.addNewTextParagraph();
        newPg.getXmlObject().set(paragraph.getXmlObject().copy());
        return newPg;
    }

    /**
     * replace with this.replaceChartData(...)
     *
     * @param chart      target chart
     * @param chartTitle title
     * @param series     series-data first row titles, B1,C1,D1 .... (skip fist, A1)
     * @param categories categories data, first column, A2,A3,A4,A5 (skip first, A1)
     * @param values     value[series.length][categories.length]
     */
    @Deprecated
    public static void setRadarData(XSLFChart chart, String chartTitle, String[] series, String[] categories, Double[]... values) {
        replaceChartData(chart, chartTitle, series, categories, values);
    }

    /**
     * data not overflow the demo PPT series/categories is batter. if overflow, PPT data will not edit by EXCEL.
     * demo PPT series size must &gt;= data.series size , this is a restricted !
     *
     * @param chart      XSLFChart to edit
     * @param chartTitle chart title
     * @param series     series-data first row titles, B1,C1,D1 .... (skip fist, A1)
     * @param categories categories data, first column, A2,A3,A4,A5 (skip first, A1)
     * @param values     value[series.length][categories.length]
     */
    public static void replaceChartData(XSLFChart chart, String chartTitle, String[] series, String[] categories, Double[]... values) {
        int size = categories.length;
        List<XDDFChartData> s = chart.getChartSeries();
        XDDFChartData bar = s.get(0);

        // THIS IS A trouble
        Assert.isTrue(bar.getSeriesCount() >= series.length, "ERROR ! Template series count >= replace.data.series count, THIS IS A restricted");

        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, size, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(//
                categories, categoryDataRange, 0);

        Assert.isTrue(series.length == values.length, "Error : series.length != values.length");

        for (int i = 0; i < series.length; i++) {
            int column = i + 1;
            String valuesDataRange = chart.formatRange(new CellRangeAddress(1, size, column, column));
            XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values[i], valuesDataRange, column);
            XDDFChartData.Series ss;
            if (bar.getSeriesCount() > i) {
                ss = bar.getSeries(i);
                ss.replaceData(categoriesData, valuesData);
            } else {
                ss = bar.addSeries(categoriesData, valuesData);
            }
            ss.setTitle(series[i], chart.setSheetTitle(series[i], column));
        }

        //DELETE overflow series
        for (int i = bar.getSeriesCount(); i > series.length; i--) {
            bar.removeSeries(i - 1);
        }

        replaceExcel(chart, series, categories, values);

        chart.plot(bar);
        chart.setTitleText(chartTitle); // https://stackoverflow.com/questions/30532612
    }

    private static void replaceExcel(XSLFChart chart, String[] series, String[] categories, Double[][] values) {
        try { //  replace workbook data . if over max row / column, PPT not edit by EXCEL, AND unsupported to add row/column
            ByteArrayOutputStream stream = new ByteArrayOutputStream();
            chart.getWorkbook().write(stream);
            XSSFWorkbook workbook = (XSSFWorkbook) XSSFWorkbookFactory.create(new ByteArrayInputStream(stream.toByteArray()));
            chart.setWorkbook(workbook);
            XSSFSheet sheet = chart.getWorkbook().getSheetAt(0);
            Assert.isTrue(sheet != null, "Must has sheet ! Template PPTX can be EDIT by micro-soft-office.");
            XSSFRow head = sheet.getRow(0);
            Assert.isTrue(head != null, "SHEET row 0 must have DATA .");
            for (int i = 0; i < series.length; i++) {
                XSSFCell cell = head.getCell(i + 1);
                if (cell != null) {
                    cell.setCellValue(series[i]);
                }
            }
            for (int i = 0; i < categories.length; i++) {
                XSSFRow row = sheet.getRow(i + 1);
                if (row == null) {
                    row = sheet.createRow(i + 1);
                }
                XSSFCell first = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (first != null) {
                    first.setCellValue(categories[i]);
                }
                for (int j = 0; j < series.length; j++) {
                    XSSFCell cell = row.getCell(j + 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if (cell != null) {
                        double v = values[j][i];
                        cell.setCellValue(v);
                    }
                }
            }
            // delete overflow column
            if (false) { //  not delete column
                for (int i = 0; i <= categories.length; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (row != null) {
                        for (int j = row.getLastCellNum() - 1; j > series.length; j--) {// from end to start
                            XSSFCell cell = row.getCell(j);
                            if (cell != null) {
                                //row.removeCell(cell); // cell.setBlank();
                                log.debug("TODO:: delete OR setBlank will do not EDIT BY PPT");
                            }
                        }
                    }
                }
            }
            // delete overflow row
            for (int i = sheet.getLastRowNum() - 1; i > categories.length; i--) {
                XSSFRow row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public static void setPieDate(XSLFChart chart, String chartTitle, String[] categories, Double[] values) {
        replaceChartData(chart, chartTitle, new String[]{"PIE-SERIES"}, categories, values);
    }
}
