package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class DeleteOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(DeleteOperation.class);

    @Override
    public void parse(ParseContext context) {
        List<?> nodes = this.search(context);
        for (Object node : nodes) {
            thisInvoke("delete", context, node);
        }
    }

    public void delete(Object o) {
        logger.warn("Not support delete type : {}", o.getClass());
    }

    public void delete(ParseContext context, XSLFTableRow row) {
        XSLFTable table = getTable(context.getTemplate(), row);
        List<XSLFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            if (rows.get(i) == row) {
                table.removeRow(i);
                break;
            }
        }
    }

    public void delete(ParseContext context, XSLFTableCell cell) {
        XSLFTableRow row = getTableRow(context.getTemplate(), cell);
        List<XSLFTableCell> cells = row.getCells();
        for (int i = 0; i < cells.size(); i++) {
            if (cells.get(i) == cell) {
                row.mergeCells(i - 1, i);
                break;
            }
        }
    }

    public static XSLFTable getTable(XSLFSheet sheet, XSLFTableRow row) {
        for (XSLFShape shape : sheet.getShapes()) {
            if (shape instanceof XSLFTable) {
                XSLFTable tb = ((XSLFTable) shape);
                for (XSLFTableRow rw : tb.getRows()) {
                    if (row.equals(rw)) {
                        return tb;
                    }
                }
            }
        }
        return null;
    }

    public static XSLFTableRow getTableRow(XSLFSheet sheet, XSLFTableCell cell) {
        for (XSLFShape shape : sheet.getShapes()) {
            if (shape instanceof XSLFTable) {
                XSLFTable tb = ((XSLFTable) shape);
                for (XSLFTableRow rw : tb.getRows()) {
                    for (XSLFTableCell c : rw.getCells()) {
                        if (c.equals(cell)) {
                            return rw;
                        }
                    }
                }
            }
        }
        return null;
    }
}
