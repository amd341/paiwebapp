package org.highpoint.paiwebapp;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;


import java.io.*;
import java.util.*;

/**
 * Created by alex on 6/13/17.
 */
public class ExcelParser {

    private HSSFWorkbook hworkbook;
    private XSSFWorkbook xworkbook;
    private Workbook workbook;
    private Map<String,String> entries;
    private boolean isXLSX;

    //figuring out how these strings correlate to real colors is annoying
    private final List<String> XLS_QUESTION_COLORS = Arrays.asList("9999:3333:6666", "7070:3030:A0A0"); //purples... don't even ask
    private final String XLS_ANSWER_COLOR = "FFFF:FFFF:0"; //yellow

    private final String XLSX_QUESTION_COLOR = "FF7030A0"; //purple
    private final String XLSX_ANSWER_COLOR = "FFFFFF00"; //yellow

    public ExcelParser(InputStream input, Map<String,String> entries, boolean isXLSX) throws IOException, InvalidFormatException {
        this.isXLSX = isXLSX;
        if (isXLSX){
            xworkbook = new XSSFWorkbook(input);
            workbook = xworkbook;
            this.entries = entries;
        }
        else{
            hworkbook = new HSSFWorkbook(input);
            workbook = hworkbook;
            this.entries = entries;
        }

    }

    /**
     * gets pairs of questions and answers based on highlighted sheets
     *
     * @throws JsonProcessingException
     */
    public List<Map<String,String>> getHighlighted() throws JsonProcessingException {
        List<Map<String,String>> sections = new ArrayList<>();
        Map<String,String> section = new HashMap<>();
        String question = "no question";
        String body;
        boolean gotQuestion = false;
        if (!isXLSX) {
            int i = 0;
            while (i < hworkbook.getNumberOfSheets()) {
                HSSFSheet sheet = hworkbook.getSheetAt(i);
                int j = 0;

                while(j <= sheet.getLastRowNum()) {
                    HSSFRow row = sheet.getRow(j);
                    int k = 0;
                    if (row != null) {
                        while ( k <= row.getLastCellNum()) {
                            HSSFCell cell = row.getCell(k);
                            if (cell != null) {
                                HSSFColor color = cell.getCellStyle().getFillForegroundColorColor();
                                if (color != null && XLS_QUESTION_COLORS.contains(color.getHexString()) && !gotQuestion) {
                                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                        question = String.valueOf(cell.getNumericCellValue());
                                        gotQuestion = true;
                                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                        question = cell.getRichStringCellValue().getString();
                                        gotQuestion = true;
                                    }
                                } else if (color != null && color.getHexString().equals(XLS_ANSWER_COLOR) && gotQuestion) {
                                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                        body = String.valueOf(cell.getNumericCellValue());
                                        section.put("question", question);
                                        section.put("body", body);
                                        section.putAll(entries);
                                        sections.add(section);
                                        section = new HashMap<>();
                                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                        body = cell.getRichStringCellValue().getString();
                                        section.put("question", question);
                                        section.put("body", body);
                                        section.putAll(entries);
                                        sections.add(section);
                                        section = new HashMap<>();

                                    }
                                }
                            }
                            k++;
                        }
                        gotQuestion = false;
                    }

                    j++;
                }
                i++;
            }
        } else {
            int i = 0;

            while (i < xworkbook.getNumberOfSheets()) {
                XSSFSheet sheet = xworkbook.getSheetAt(i);
                int j = 0;
                while (j <= sheet.getLastRowNum()) {
                    XSSFRow row = sheet.getRow(j);
                    int k = 0;
                    if (row != null) {
                        while (k <= row.getLastCellNum()) {
                            XSSFCell cell = row.getCell(k);
                            if (cell != null) {
                                XSSFColor color = cell.getCellStyle().getFillForegroundColorColor();
                                if (color != null && color.getARGBHex().equals(XLSX_QUESTION_COLOR) && !gotQuestion) {
                                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                        question = String.valueOf(cell.getNumericCellValue());
                                        gotQuestion = true;
                                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                        question = cell.getRichStringCellValue().getString();
                                        gotQuestion = true;

                                    }
                                } else if (color != null && color.getARGBHex().equals(XLSX_ANSWER_COLOR) && gotQuestion) {
                                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                        body = String.valueOf(cell.getNumericCellValue());
                                        section.put("question", question);
                                        section.put("body", body);
                                        section.putAll(entries);
                                        sections.add(section);
                                        section = new HashMap<>();
                                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                        body = cell.getRichStringCellValue().getString();
                                        section.put("question", question);
                                        section.put("body", body);
                                        section.putAll(entries);
                                        sections.add(section);
                                        section = new HashMap<>();

                                    }
                                }
                            }
                            k++;
                        }
                        gotQuestion = false;
                    }

                    j++;
                }
                i++;
            }


        }
        return sections;

    }

    //I think that this method should not be an option for use. It doesn't really return meaningful data.
    public List<Map<String,String>> getSheets() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,String> sectionHash = new HashMap<>();
        List<Map<String,String>> sectionsList = new ArrayList<>();
        boolean isBlankRow = true;

        Sheet datatypeSheet = workbook.getSheetAt(0);
        StringBuilder body = new StringBuilder();
        //iterator to iterate through sheets
        //Iterator<Sheet> sheetIterator = workbook.iterator();

        int count = 0;
        while(count < workbook.getNumberOfSheets()){
            datatypeSheet = workbook.getSheetAt(count);

            sectionHash.putAll(entries);
            sectionHash.put("heading", workbook.getSheetName(count));

            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {


                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();

                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        body.append(currentCell.getStringCellValue()).append(" ");
                        isBlankRow = false;
                    }
                    else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        body.append(currentCell.getNumericCellValue()).append(" ");
                        isBlankRow = false;
                    }
                }
                if (!isBlankRow){
                    body.append("\n");
                    isBlankRow = true;
                }

            }
            count++;
            sectionHash.put("body",body.toString());
            sectionsList.add(sectionHash);
            sectionHash = new HashMap<>();
            body = new StringBuilder();
        }

        return(sectionsList);
    }

    public List<String> getJsonStrings() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,String> sectionHash = new HashMap<>();
        List<String> sectionsList = new ArrayList<>();
        boolean isBlankRow = true;

        Sheet datatypeSheet = workbook.getSheetAt(0);
        StringBuilder body = new StringBuilder();

        int count = 0;
        while(count < workbook.getNumberOfSheets()){
            datatypeSheet = workbook.getSheetAt(count);

            sectionHash.putAll(entries);
            sectionHash.put("heading", workbook.getSheetName(count));

            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {


                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();

                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        body.append(currentCell.getStringCellValue()).append("--"); //should the -- be removed?
                        isBlankRow = false;
                    }
                    else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        body.append(currentCell.getNumericCellValue()).append("--"); //same as above
                        isBlankRow = false;
                    }
                }
                if (!isBlankRow){
                    body.append("\n");
                    isBlankRow = true;
                }

            }
            count++;
            sectionHash.put("body",body.toString());
            sectionsList.add(objectMapper.writeValueAsString(sectionHash));
            sectionHash = new HashMap<>();
            body = new StringBuilder();
        }

        return(sectionsList);
    }






}