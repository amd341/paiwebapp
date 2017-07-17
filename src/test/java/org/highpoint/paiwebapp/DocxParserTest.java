package org.highpoint.paiwebapp;

import junit.framework.TestCase;

import java.util.*;


/**
 * Created by brenden on 7/13/17.
 * only tests the two methods that are actually used by the application, as the others are mostly useful for testing
 * /really aren't needed anymore, but kept around because why not
 */
public class DocxParserTest extends TestCase {
    public void testGetHighlighted() throws Exception {
        Map<String,String> map = new HashMap<>();
        map.put("test","test");

        DocxParser p = new DocxParser(getClass().getResourceAsStream("/testh.docx"), map);
        List<Map<String,String>> sections = p.getHighlighted();

        Map<String,String> expectedSection1 = new HashMap<>();
        Map<String,String> expectedSection2 = new HashMap<>();

        expectedSection1.put("question", "Test1");
        expectedSection1.put("body", "Test1");
        expectedSection1.put("test","test");

        expectedSection2.put("question","Test2");
        expectedSection2.put("body","Test2");
        expectedSection2.put("test","test");


        List<Map<String,String>> expectedSections = new ArrayList<>(Arrays.asList(expectedSection1,expectedSection2));
        assertEquals(expectedSections, sections);


    }

    /*this test may give some insight about when Word will add newlines. Although for our purposes, the newlines
    * are inconsequential, it may still be useful to know, or at least have an idea about this, as it's hidden elsewhere.
    * From what I've seen, one is added if you're right before a heading, and a bunch are added with tables/end of document? */
    public void testGetSubsections() throws Exception {
        Map<String,String> map = new HashMap<>();
        map.put("test","test");

        DocxParser p = new DocxParser(getClass().getResourceAsStream("/test.docx"), map);
        List<Map<String,String>> sections = p.getSubSections();

        List<Map<String,String>> expectedSections = new ArrayList<>();

        Map<String,String> expectedSection = new HashMap<>();

        expectedSection.put("headingOne", "H1");
        expectedSection.put("headingTwo", "H2");
        expectedSection.put("body","Text\n");
        expectedSection.put("test","test");

        expectedSections.add(expectedSection);
        expectedSection = new HashMap<>();

        expectedSection.put("headingOne", "H1");
        expectedSection.put("headingTwo", "H2");
        expectedSection.put("test","test");
        expectedSection.put("body","More\n\n\n");

        expectedSections.add(expectedSection);

        assertEquals(expectedSections, sections);



    }

}