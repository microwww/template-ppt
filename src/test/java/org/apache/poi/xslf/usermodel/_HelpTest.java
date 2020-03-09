package org.apache.poi.xslf.usermodel;

import com.github.microwww.ttp.Tools;
import com.github.microwww.ttp.util.BiConsumer;
import com.github.microwww.ttp.util._Help;
import com.github.microwww.ttp.xslf.XSLFGraphicChart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.UUID;

public class _HelpTest {

    public static final String PATH = _HelpTest.class.getResource("/").getFile();

    @Test
    public void copyTable() throws IOException {
        XMLSlideShow target, template;
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }
        XSLFSlide slide = target.createSlide();
        XSLFTable shape = (XSLFTable) template.getSlides().get(0).getShapes().get(0);
        // 1
        XSLFTable table = _Help.copyTable(slide, shape);
        Assert.assertEquals(table.getRows().size(), shape.getRows().size());
        Assert.assertEquals(table.getRows().get(0).getCells().size(), shape.getRows().get(0).getCells().size());

        // 2
        XSLFTableRow orow = shape.getRows().get(1);
        XSLFTableRow row = _Help.copyTableRow(table, orow);
        Assert.assertEquals(row.getCells().size(), orow.getCells().size());

        target.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }

    @Test
    public void copyChart() throws IOException, InvalidFormatException {
        XMLSlideShow target, template;
        try (FileInputStream in = new FileInputStream(new File(PATH, "chart.pptx"))) {
            template = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(PATH, "chart.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }

        target.createSlide().importContent(template.getSlides().get(0));

        final XSLFSlide slide = target.createSlide();
        XSLFChart chart = _Help.copyChart(template.getSlides().get(0), 0, slide, new BiConsumer<XSLFChart, XSLFGraphicChart>() {
            public void accept(XSLFChart c, XSLFGraphicChart val) {
                slide.addChart(c, _Help.delta(val.getGraphic().getAnchor(), 0, 0, 0, 0));
            }
        });
        Assert.assertNotNull(chart);

        target.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }

    @Test
    public void setChartData() throws IOException {
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(PATH, "chart.pptx"))) {
            template = new XMLSlideShow(in);
        }

        XSLFGraphicChart position = _Help.findChartWithPosition(template.getSlides().get(1), 0);
        logChartData(position);

        //XSLFGraphicChart position = _Help.findChartWithPosition(template.getSlides().get(0), 0);
        Tools.setRadarData(position.getChart(), "title", new String[]{"s1", "s2", "s3"}, new String[]{"c1", "c2", "c3", "c4"},
                new Double[]{7.1, 6.1, 5.1, 4.1}, new Double[]{6.2, 5.2, 4.2, 3.2}, new Double[]{5.3, 4.3, 3.3, 2.3});
        //Tools.setRadarData(position.getChart(), "title", new String[]{"s1", "s2", "s3", "s4"}, new String[]{"c1", "c2", "c3", "c4", "c5"},
        //        new Double[]{7.1, 6.1, 5.1, 4.1, 3.1}, new Double[]{6.2, 5.2, 4.2, 3.2, 2.2}, new Double[]{5.3, 4.3, 3.3, 2.3, 1.3},
        //        new Double[]{4.4, 3.4, 2.4, 1.4, 0.4});
        //Tools.setRadarData(position.getChart(), "title", new String[]{"s1", "s2"}, new String[]{"c1", "c2", "c3"},
        //        new Double[]{7.1, 6.1, 5.1}, new Double[]{6.2, 5.2, 4.2});

        template.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }

    public void logChartData(XSLFGraphicChart chart) {
        XSLFChart origin = chart.getChart();
        List<XDDFChartData> ss = origin.getChartSeries();
        for (int k = 0; k < ss.size(); k++) {
            XDDFChartData s = ss.get(k);
            for (int i = 0; i < s.getSeriesCount(); i++) {
                XDDFNumericalDataSource<? extends Number> vd = s.getSeries(i).getValuesData();
                XDDFDataSource<?> data = s.getSeries(i).getCategoryData();
                for (int j = 0; j < data.getPointCount(); j++) {
                    System.out.print(vd.getPointAt(j) + ", ");
                }
                System.out.println();
            }
        }
    }
}