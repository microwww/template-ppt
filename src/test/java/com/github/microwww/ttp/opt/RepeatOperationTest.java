package com.github.microwww.ttp.opt;

import org.junit.Test;

import java.util.List;

import static org.junit.Assert.*;

public class RepeatOperationTest {

    @Test
    public void getElement() {
    }

    @Test
    public void searchRanges() {
        {
            List<Range> ranges = Operation.searchRanges("1,5");
            assertEquals(ranges.size(), 2);
            assertEquals(ranges.get(0).getFrom(), 1);
            assertEquals(ranges.get(0).getTo(), 2);
            assertEquals(ranges.get(1).getFrom(), 5);
            assertEquals(ranges.get(1).getTo(), 6);
        }
        {
            List<Range> ranges = Operation.searchRanges("1-");
            assertEquals(ranges.size(), 1);
            assertEquals(ranges.get(0).getFrom(), 1);
            assertEquals(ranges.get(0).getTo(), Integer.MAX_VALUE);
        }
        {
            List<Range> ranges = Operation.searchRanges("2,5-8,11");
            assertEquals(ranges.size(), 3);
            assertEquals(ranges.get(0).getFrom(), 2);
            assertEquals(ranges.get(0).getTo(), 3);
            assertEquals(ranges.get(1).getFrom(), 5);
            assertEquals(ranges.get(1).getTo(), 8);
            assertEquals(ranges.get(2).getFrom(), 11);
            assertEquals(ranges.get(2).getTo(), 12);
        }
        {
            List<Range> ranges = Operation.searchRanges("2,5-8,11-20,30-");
            assertEquals(ranges.size(), 4);
            assertEquals(ranges.get(0).getFrom(), 2);
            assertEquals(ranges.get(0).getTo(), 3);
            assertEquals(ranges.get(1).getFrom(), 5);
            assertEquals(ranges.get(1).getTo(), 8);
            assertEquals(ranges.get(2).getFrom(), 11);
            assertEquals(ranges.get(2).getTo(), 20);
            assertEquals(ranges.get(3).getFrom(), 30);
            assertEquals(ranges.get(3).getTo(), Integer.MAX_VALUE);
        }

    }
}