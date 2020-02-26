package com.github.microwww.ttp.opt;

import org.junit.Test;

import java.io.File;
import java.io.IOException;

import static org.junit.Assert.assertEquals;

public class ParseExpressesTest {

    @Test
    public void parse() throws IOException {
        ParseExpresses exp = new ParseExpresses(new File(this.getClass().getResource("/").getFile(), "demo.txt"));
        exp.parse();
        assertEquals(8, exp.getOperations().size());
        assertEquals(2, exp.getOperations().get(5).childrenOperations.size());
        assertEquals(1, exp.getOperations().get(5).childrenOperations.get(1).childrenOperations.size());
    }

}