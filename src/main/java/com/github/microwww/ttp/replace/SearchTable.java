package com.github.microwww.ttp.replace;

import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableRow;

import java.util.ArrayList;
import java.util.List;

public class SearchTable implements SearchContent {
        private final XSLFTable table;

        public SearchTable(XSLFTable table) {
            this.table = table;
        }

        @Override
        public List<TextExpress> search() {
            ArrayList<TextExpress> list = new ArrayList<>();
            for(XSLFTableRow row : table.getRows()){
                List<TextExpress> search = new SearchTableRow(row).search();
                list.addAll(search);
            }
            return list;
        }
    }