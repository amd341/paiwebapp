package org.highpoint.paiwebapp;

import junit.framework.TestCase;

import java.util.*;

/**
 * Created by brenden on 7/17/17.
 */
public class ExcelParserTest extends TestCase {
    public void testGetHighlightedJson() throws Exception {
        Map<String,String> map = new HashMap<>();
        map.put("test","test");

        ExcelParser p1 = new ExcelParser(getClass().getResourceAsStream("/test.xlsx"), map, true);
        ExcelParser p2 = new ExcelParser(getClass().getResourceAsStream("/test.xls"), map, false);

        List<Map<String,String>> sectionsOne = p1.getHighlighted();
        List<Map<String,String>> sectionsTwo = p2.getHighlighted();

        List<Map<String,String>> expectedSections = new ArrayList<>();

        Map<String,String> expectedSection = new HashMap<>();

        expectedSection.put("test", "test");
        expectedSection.put("question","Q1");
        expectedSection.put("body","A1");

        expectedSections.add(expectedSection);
        expectedSection = new HashMap<>();

        expectedSection.put("test", "test");
        expectedSection.put("question","Q2");
        expectedSection.put("body","A2a");

        expectedSections.add(expectedSection);
        expectedSection = new HashMap<>();

        expectedSection.put("test", "test");
        expectedSection.put("question","Q2");
        expectedSection.put("body","A2b");

        expectedSections.add(expectedSection);
        expectedSection = new HashMap<>();

        expectedSection.put("test", "test");
        expectedSection.put("question","Q3");
        expectedSection.put("body","A3");

        expectedSections.add(expectedSection);

        assertEquals(expectedSections, sectionsOne);
        assertEquals(expectedSections, sectionsTwo);



    }

}