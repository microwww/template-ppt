package org.apache.poi.xslf.usermodel;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
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
import java.util.AbstractMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

public class _Help {

    public static XSLFTable copyTable(XSLFSlide slide, XSLFTable src) {
        XSLFTable dest = slide.createTable();
        dest.getCTTable().set(src.getCTTable().copy());

        List<CTTableRow> tr = dest.getCTTable().getTrList();
        for (CTTableRow row : tr) {
            try {
                List<XSLFTableRow> rows = (List<XSLFTableRow>)
                        FieldUtils.readDeclaredField(dest, "_rows", true);
                rows.add(new XSLFTableRow(row, dest));
            } catch (IllegalAccessException e) {
                throw new RuntimeException("error! this invoke private method ...", e);
            }
            dest.updateRowColIndexes();
        }
        dest.copy(src);
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
                _cells.add(new XSLFTableCell(tc, table));
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }
        return row;
    }

    public static XSLFChart copyChart(XSLFSlide src, int index, XSLFSlide dest) {
        return copyChart(src, index, dest, (c, v) -> {
            dest.addChart(c, rectanglePx2point(v.getKey().getAnchor(), 0, 0, 0, 0));
        });
    }

    public static XSLFChart copyChart(XSLFSlide src, int index, XSLFSlide dest, BiConsumer<XSLFChart, Map.Entry<XSLFGraphicFrame, XSLFChart>> fun) {
        Map.Entry<XSLFGraphicFrame, XSLFChart> position = findChartWithPosition(src, index);
        if (position != null) {
            XSLFChart from = position.getValue();
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

    // copy from  XSLFGraphicFrame.copy
    public static Map.Entry<XSLFGraphicFrame, XSLFChart> findChartWithPosition(XSLFSlide src, int index) {
        int i = 0;
        for (XSLFShape sh : src.getShapes()) {
            if (sh instanceof XSLFGraphicFrame) {
                XSLFGraphicFrame frame = (XSLFGraphicFrame) sh;
                CTGraphicalObjectData data = ((CTGraphicalObjectFrame) frame.getXmlObject()).getGraphic().getGraphicData();
                String uri = data.getUri();
                if (uri.endsWith("/chart")) {
                    if (i == index) {
                        XSLFChart chart = findChart(frame);
                        return new AbstractMap.SimpleEntry<>(frame, chart);
                    }
                    i++;
                }
            }
        }
        return null;
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
