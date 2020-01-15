package com.github.microwww.ttp;

public class Assert {
    public static void isTrue(boolean val, String msg) {
        if (!val) {
            throw new IllegalArgumentException(msg);
        }
    }

    public static void isNotNull(Object val, String msg) {
        if (val == null) {
            throw new IllegalArgumentException(msg);
        }
    }
}
