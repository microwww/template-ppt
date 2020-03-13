package com.github.microwww.ttp.opt;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ElseOperation extends Operation {

    private static final Logger logger = LoggerFactory.getLogger(ElseOperation.class);

    @Override
    public void parse(ParseContext context) {
        throw new UnsupportedOperationException("This is a FLAG, only as IfOperation children");
    }
}
