package com.github.microwww.ttp.operation;

public interface ElementsOperation<T, C> {
    void copy(T src, C shapes);

    void delete(T src, C shapes);

    void replace(T src, C shapes);
}
