package com.github.microwww.ttp.opt;

import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.junit.Assert.assertEquals;

public class ParseExpressesTest {

    @Test
    public void parse() throws IOException {
        ParseExpresses parser = new ParseExpresses();
        List<Operation> exp = parser.parse(new File(this.getClass().getResource("/").getFile(), "demo.txt"));
        assertEquals(8, exp.size());
        assertEquals(2, exp.get(5).childrenOperations.size());
    }

}