package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParseExpresses {

    private final InputStream inputStream;
    private final List<Operation> operations = new ArrayList<>();

    public ParseExpresses(File file) throws FileNotFoundException {
        this.inputStream = new FileInputStream(file);
    }

    public ParseExpresses(InputStream inputStream) throws FileNotFoundException {
        this.inputStream = inputStream;
    }

    public List<Operation> getOperations() {
        return operations;
    }

    public void parse() throws IOException {
        BufferedReader in = new BufferedReader(new InputStreamReader(inputStream, "UTF-8"));
        while (true) {
            String ln = in.readLine();
            if (ln == null) {
                break;
            }
            ln = ln.trim();
            if (ln.length() == 0 || ln.startsWith("#")) {
                continue;
            }
            String[] exps, params = new String[]{};
            if (ln.endsWith(")")) {
                int idx = ln.indexOf('(', 0);
                String[] opts = new String[]{ln.substring(0, idx), ln.substring(idx + 1)};
                // Assert.isTrue(opts.length == 2, "if has params , must has '(',')' , and end with ')'");
                exps = opts[0].split(" +");
                Assert.isTrue(exps.length > 0, "must has express");
                params = opts[1].substring(0, opts[1].length() - 1).trim().split(" +");
            } else {
                exps = ln.split(" +");
            }
            Operation operation = toOptions(exps, params);
            operations.add(operation);
        }
    }

    public Operation toOptions(String[] exps, String[] params) {
        String prefix = null;
        for (int i = 0; i < exps.length; i++) {
            String ex = exps[i];
            Matcher matcher = Pattern.compile("[a-zA-Z]+").matcher(ex);
            if (matcher.matches()) {
                String exp = ex.substring(0, 1).toUpperCase() + ex.substring(1);
                try {
                    String name = this.getClass().getPackage().getName() + "." + exp + "Operation";
                    Class<? extends Operation> clazz = (Class<? extends Operation>) Class.forName(name);
                    Operation operation = clazz.getConstructor().newInstance();
                    operation.setPrefix(prefix);
                    operation.setNode(ArrayUtils.subarray(exps, i + 1, exps.length));
                    operation.setParams(params);
                    return operation;
                } catch (ClassNotFoundException | NoSuchMethodException | InstantiationException | IllegalAccessException | InvocationTargetException e) {
                    throw new UnsupportedOperationException("Not support type :: " + ex, e);
                }
            } else {
                if (prefix == null) {
                    prefix = ex;
                }
            }
        }
        throw new UnsupportedOperationException(String.format("Not support type :: %s ( %s )",
                StringUtils.join(exps, " "), StringUtils.join(params, " ")));
    }
}
