package com.github.microwww.ttp.opt;

public class RepeatDomain {
    private Object item;
    private int index;
    private Object parent;

    public Object getItem() {
        return item;
    }

    public RepeatDomain setItem(Object item) {
        this.item = item;
        return this;
    }

    public int getIndex() {
        return index;
    }

    public RepeatDomain setIndex(int index) {
        this.index = index;
        return this;
    }

    public Object getParent() {
        return parent;
    }

    public void setParent(Object parent) {
        this.parent = parent;
    }
}