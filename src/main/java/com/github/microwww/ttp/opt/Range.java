package com.github.microwww.ttp.opt;

public class Range {
    private final int from;
    private final int to;

    public Range(int from, int to) {
        this.from = from;
        this.to = to;
    }

    public int getFrom() {
        return from;
    }

    public int getTo() {
        return to;
    }

    public boolean isIn(int idx) {
        return idx >= from && idx < to;
    }
}