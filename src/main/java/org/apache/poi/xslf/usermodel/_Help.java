package org.apache.poi.xslf.usermodel;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableCell;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableRow;

import java.util.List;

public class _Help {

    public static XSLFTable copeTable(XSLFSlide slide, XSLFTable src) {
        XSLFTable dest = slide.createTable();
        dest.getCTTable().set(src.getCTTable().copy());
        slide.getShapes(); // this.initDrawingAndShapes();
        dest.setAnchor(src.getAnchor());

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
}
