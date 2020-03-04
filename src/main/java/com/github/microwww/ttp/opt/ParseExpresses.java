package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParseExpresses {

    private final List<Class<? extends Operation>> supportOperations = new CopyOnWriteArrayList<>();

    public ParseExpresses addSupportOperations(Class<? extends Operation>... supportOperations) {
        this.supportOperations.addAll(Arrays.asList(supportOperations));
        return this;
    }

    public List<Operation> parse(File input) throws IOException {
        return this.parse(new InputStreamReader(new FileInputStream(input), "UTF-8"));
    }

    public List<Operation> parse(InputStream input) throws IOException {
        return this.parse(new InputStreamReader(input, "UTF-8"));
    }

    public List<Operation> parse(InputStreamReader input) throws IOException {
        List<Operation> operations = new ArrayList<>();
        BufferedReader in = new BufferedReader(input);
        Operation upper = null;
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
            String fix = operation.getPrefix();
            if (fix != null && fix.startsWith(">")) {
                Assert.isTrue(upper != null, "Express '>' NOT at first");
                int uplen = 0;
                if (upper.getPrefix() != null) {
                    uplen = upper.getPrefix().length();
                }
                int len = operation.getPrefix().length();
                Assert.isTrue(len <= 1 + uplen, "Express '>' count must big-equal Up+1 COUNT");
                for (int i = 0; i <= uplen - len; i++) {
                    upper = upper.getParentOperations();
                }
                operation.setParentOperations(upper);
                upper.addChildrenOperation(operation);
            } else {
                operations.add(operation);
            }
            upper = operation;
        }
        return operations;
    }

    public Operation toOptions(String[] exps, String[] params) {
        String prefix = null;
        for (int i = 0; i < exps.length; i++) {
            String ex = exps[i];
            Matcher matcher = Pattern.compile("[a-zA-Z]+").matcher(ex);
            if (matcher.matches()) {
                String exp = ex.substring(0, 1).toUpperCase() + ex.substring(1);
                try {
                    Class<? extends Operation> clazz = parseOperationClass(exp);
                    Operation operation = clazz.getConstructor().newInstance();
                    operation.setPrefix(prefix);
                    operation.setNode(ArrayUtils.subarray(exps, i + 1, exps.length));
                    operation.setParams(params);
                    return operation;
                } catch (NoSuchMethodException | InstantiationException | IllegalAccessException | InvocationTargetException e) {
                    throw new UnsupportedOperationException("Parse " + exp + " error. Operation class MUST has a no-params-constructor", e);
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

    protected Class<? extends Operation> parseOperationClass(String exp) {
        Class<? extends Operation> clazz = null;
        for (Class<? extends Operation> opt : this.supportOperations) {
            if (opt.getSimpleName().equals(exp + "Operation")) {
                clazz = opt;
            }
        }
        if (clazz == null) {
            String name = this.getClass().getPackage().getName() + "." + exp + "Operation";
            try {
                clazz = (Class<? extends Operation>) Class.forName(name);
            } catch (ClassNotFoundException e) {
                throw new UnsupportedOperationException("Not support type :: " + exp + ", Type name must end with 'Operation'", e);
            }
        }
        return clazz;
    }
}
