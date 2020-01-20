package com.github.microwww.ttp.opt;

import org.junit.Test;

import java.io.File;
import java.io.IOException;

import static org.junit.Assert.*;

public class ParseExpressesTest {

    @Test
    public void parse() throws IOException {
        ParseExpresses exp = new ParseExpresses(new File(this.getClass().getResource("/").getFile(), "demo.txt"));
        exp.parse();
        assertEquals(14, exp.getOperations().size());
    }
}