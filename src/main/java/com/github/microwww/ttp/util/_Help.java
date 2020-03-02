package com.github.microwww.ttp.util;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.reflect.MethodUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableCell;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;

import javax.xml.namespace.QName;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class _Help {

    public static XSLFTable copyTable(XSLFSheet sheet, XSLFTable src) {
        XSLFTable dest = sheet.createTable();
        dest.getCTTable().set(src.getCTTable().copy());

        List<CTTableRow> tr = dest.getCTTable().getTrList();
        try {
            for (CTTableRow row : tr) {
                List<XSLFTableRow> rows = (List<XSLFTableRow>)
                        FieldUtils.readDeclaredField(dest, "_rows", true);
                Constructor<XSLFTableRow> constructor = XSLFTableRow.class.getDeclaredConstructor(CTTableRow.class, XSLFTable.class);
                constructor.setAccessible(true);
                rows.add(constructor.newInstance(row, dest));
                MethodUtils.invokeMethod(dest, true, "updateRowColIndexes");
            }
            MethodUtils.invokeMethod(dest, true, "copy", src);
        } catch (IllegalAccessException | NoSuchMethodException | InstantiationException | InvocationTargetException e) {
            throw new RuntimeException("error! this invoke private method ...", e);
        }
        return dest;
    }

    public static XSLFTableRow copyTableRow(XSLFTable table, XSLFTableRow src) {
        XSLFTableRow row = table.addRow();
        row.getXmlObject().set(src.getXmlObject().copy());
        for (CTTableCell tc : row.getXmlObject().getTcList()) {
            //XSLFTableRow._cells.add(new XSLFTableCell(cell, table));
            try {
                List<XSLFTableCell> _cells = (List<XSLFTableCell>)
                        FieldUtils.readDeclaredField(row, "_cells", true);
                Constructor<XSLFTableCell> constructor = XSLFTableCell.class.getDeclaredConstructor(CTTableCell.class, XSLFTable.class);
                constructor.setAccessible(true);
                _cells.add(constructor.newInstance(tc, table));
            } catch (IllegalAccessException | InstantiationException | InvocationTargetException | NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        }
        return row;
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, final XSLFSheet dest) {
        return copyChart(src, index, dest, new BiConsumer<XSLFChart, XSLFGraphicChart>() {
            @Override
            public void accept(XSLFChart c, XSLFGraphicChart v) {
                Rectangle point = rectanglePx2point(v.getGraphic().getAnchor(), 0, 0, 0, 0);
                dest.addChart(c, point);
            }
        });
    }

    public static XSLFChart copyChart(XSLFSheet src, int index, XSLFSheet dest, BiConsumer<XSLFChart, XSLFGraphicChart> fun) {
        XSLFGraphicChart position = findChartWithPosition(src, index);
        if (position != null) {
            XSLFChart from = position.getChart();
            XSLFChart to = dest.getSlideShow().createChart();
            fun.accept(to, position);
            //dest.addChart(to, rectanglePx2point(position.getKey().getAnchor(), delta.x, delta.y, delta.width, delta.height));
            try {// https://github.com/apache/poi/blob/bb2ad49a2fc6c74948f8bb92701807093b525586/src/examples/src/org/apache/poi/xslf/usermodel/ChartFromScratch.java
                to.importContent(from);
                to.setWorkbook(from.getWorkbook());
            } catch (IOException | InvalidFormatException e) {
                throw new RuntimeException(e);
            }
            return to;
        }
        return null;
    }

    /**
     * @param slide
     * @param index from 0
     * @return maybe null
     */
    public static XSLFChart findChart(XSLFSheet slide, int index) {
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

    // copy from  XSLFGraphicFrame.copy
    public static XSLFGraphicChart findChartWithPosition(XSLFSheet src, int index) {
        List<XSLFGraphicChart> inx = listCharts(src);
        if (inx.size() > index && index >= 0) {
            return inx.get(index);
        }
        return null;
    }

    public static List<XSLFGraphicChart> listCharts(XSLFSheet src) {
        List<XSLFGraphicChart> list = new ArrayList<>();
        for (XSLFShape sh : src.getShapes()) {
            if (sh instanceof XSLFGraphicFrame) {
                XSLFGraphicFrame frame = (XSLFGraphicFrame) sh;
                CTGraphicalObjectData data = ((CTGraphicalObjectFrame) frame.getXmlObject()).getGraphic().getGraphicData();
                String uri = data.getUri();
                if (uri.endsWith("/chart")) {
                    XSLFChart chart = findChart(frame);
                    list.add(new XSLFGraphicChart(frame, chart));
                }
            }
        }
        return list;
    }

    // copy from  XSLFGraphicFrame.copy
    public static XSLFChart findChart(XSLFGraphicFrame srcShape) {
        CTGraphicalObjectData objData = ((CTGraphicalObjectFrame) srcShape.getXmlObject()).getGraphic().getGraphicData();
        XSLFSheet src = srcShape.getSheet();
        String xpath = "declare namespace c='http://schemas.openxmlformats.org/drawingml/2006/chart' c:chart";
        XmlObject[] obj = objData.selectPath(xpath);
        if (obj != null && obj.length == 1) {
            XmlCursor c = obj[0].newCursor();
            try {
                QName idQualifiedName = new QName("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                String id = c.getAttributeText(idQualifiedName);
                return (XSLFChart) src.getRelationById(id);
            } finally {
                c.dispose();
            }
        }
        throw new RuntimeException("Not find chart AT xpath : " + xpath);
    }

    public static Rectangle rectanglePx2point(Rectangle2D px, double x, double y, double w, double h) {
        return new Rectangle(Units.toEMU(px.getX() + x), Units.toEMU(px.getY() + y), Units.toEMU(px.getWidth() + w), Units.toEMU(px.getHeight() + h));
    }
}
