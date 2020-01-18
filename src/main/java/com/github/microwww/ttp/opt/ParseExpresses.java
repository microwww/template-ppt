package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

public class ParseExpresses {
    File file;
    List<Operation> operations = new ArrayList<>();

    public ParseExpresses(File file) {
        this.file = file;
    }

    public void parse() throws IOException {
        BufferedReader in = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
        while (true) {
            String ln = in.readLine();
            if (ln == null) {
                break;
            }
            ln = ln.trim();
            if (ln.length() == 0) {
                continue;
            }
            String[] exps;
            String[] params = new String[]{};
            if (ln.endsWith(")")) {
                String[] opts = StringUtils.split(ln, '(');
                Assert.isTrue(opts.length == 2, "if has params , must has '(',')' , and end with ')'");
                exps = opts[0].split(" +");
                Assert.isTrue(exps.length > 0, "must has express");
                params = opts[1].substring(0, opts[1].length() - 1).trim().split(" +");
            } else {
                exps = ln.split(" +");
            }
            Operation ot = new Operation();
            ot.setExpresses(exps);
            ot.setParams(params);
            operations.add(ot);
        }
    }
}
