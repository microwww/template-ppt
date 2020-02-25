package com.github.microwww.ttp.util;

import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public final class DataUtil {

    private DataUtil() {
    }

    public static List toList(Object o) {
        if (o instanceof List) {
            return (List) o;
        } else if (o instanceof Collection) {
            return new ArrayList<>((Collection) o);
        } else if (o.getClass().isArray()) {
            int size = Array.getLength(o);
            List list = new ArrayList<>();
            for (int i = 0; i < size; i++) {
                list.add(Array.get(o, i));
            }
            return list;
        }
        throw new UnsupportedOperationException("Result must is ARRAY / LIST / COLLECTION");
    }
}
