package org.highpoint.paiwebapp;

import junit.framework.TestCase;
import org.hamcrest.collection.IsIterableContainingInOrder;
import org.junit.Assert;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

import static org.hamcrest.CoreMatchers.is;

/**
 * Created by brenden on 7/13/17.
 */
public class DocxParserTest extends TestCase {
    public void testGetHighlighted() throws Exception {
        Map<String,Object> map = new HashMap<>();
        map.put("test","test");

        DocxParser p = new DocxParser(getClass().getResourceAsStream("/testh.docx"), map);
        List<Map<String,Object>> sections = p.getHighlighted();

        Map<String,Object> expectedSection1 = new HashMap<>();
        Map<String,Object> expectedSection2 = new HashMap<>();

        expectedSection1.put("question", "Test1");
        expectedSection1.put("body", "Test1");
        expectedSection1.put("test","test");

        expectedSection2.put("question","Test2");
        expectedSection2.put("body","Test2");
        expectedSection2.put("test","test");


        List<Map<String,Object>> expectedSections = new ArrayList<>(Arrays.asList(expectedSection1,expectedSection2));
        assertEquals(expectedSections, sections);


    }

    public void testGetSubSections() throws Exception {

        //heading, then paragraph, then heading, then table
        Map<String,Object> map = new HashMap<>();
        map.put("test", "test");
        DocxParser p = new DocxParser(getClass().getResourceAsStream("/test.docx"), map);
        List<Map<String,Object>> sections = p.getSections();

        Map<String,Object> expectedSection1 = new HashMap<>();
        Map<String,Object> expectedSection2 = new HashMap<>();
        expectedSection1.put("heading","Heading 1");
        expectedSection1.put("body", "Paragraph\n");
        expectedSection1.put("test", "test");
        expectedSection2.put("heading", "Heading 1");
        expectedSection2.put("body", "Table\n\n\n");
        expectedSection2.put("test","test");
        List<Map<String,Object>> expectedSections = new ArrayList<>(Arrays.asList(expectedSection1,expectedSection2));

        assertEquals(expectedSections, sections);

        //heading, then table, then heading, then paragraph
        p = new DocxParser(getClass().getResourceAsStream("/test1.docx"), map);
        sections = p.getSections();
        expectedSection2.put("body", "Table\n\n");
        expectedSections = new ArrayList<>(Arrays.asList(expectedSection2,expectedSection1));

        assertEquals(expectedSections, sections);

        //only paragraph
        p = new DocxParser(getClass().getResourceAsStream("/test2.docx"), map);
        sections = p.getSections();
        expectedSection1.put("heading", "No heading");
        expectedSections = new ArrayList<>(Collections.singletonList(expectedSection1));

        assertEquals(expectedSections, sections);

        //only table
        p = new DocxParser(getClass().getResourceAsStream("/test3.docx"), map);
        sections = p.getSections();

        expectedSection2.put("body", "Table\n\n\n"); //newlines are hard
        expectedSection2.put("heading", "No heading");
        expectedSections = new ArrayList<>(Collections.singletonList(expectedSection2));

        assertEquals(expectedSections, sections);
    }

}