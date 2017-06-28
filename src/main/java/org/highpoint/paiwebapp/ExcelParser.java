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
    private Map<String,Object> entries;
    private boolean isXLSX;

    private final String XLS_QUESTION_COLOR = "9999:3333:6666";
    private final String XLS_ANSWER_COLOR = "FFFF:FFFF:0";

    private final String XLSX_QUESTION_COLOR = "FF7030A0";
    private final String XLSX_ANSWER_COLOR = "FFFFFF00";
    //private static final String FILE_NAME = "/home/alex/documents/excels/enterprise.xlsx";

    public ExcelParser(InputStream input, Map<String,Object> entries, boolean isXLSX) throws IOException, InvalidFormatException {
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

    public List<String> getJsonStrings() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,Object> sectionHash = new HashMap<>();
        List<String> sectionsList = new ArrayList<>();
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
            sectionHash.put("body",body);
            sectionsList.add(objectMapper.writeValueAsString(sectionHash));
            sectionHash = new HashMap<>();
            body = new StringBuilder();
        }

        return(sectionsList);
    }

    public List<String> getHighlightedJsonStrings() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        List<String> sections = new ArrayList<>();
        Map<String,Object> section = new HashMap<>();
        String question = "no question";
        String body;
        boolean gotQuestion = false;
        if (!isXLSX) {
            int i = 0;

            while (i < hworkbook.getNumberOfSheets()) {
                HSSFSheet sheet = hworkbook.getSheetAt(i);
                int j = 0;

                while(j < sheet.getLastRowNum()) {
                    HSSFRow row = sheet.getRow(j);
                    int k = 0;
                    if (row != null) {
                        while ( k < row.getLastCellNum()) {
                            HSSFCell cell = row.getCell(k);
                            if (cell != null) {
                                HSSFColor color = cell.getCellStyle().getFillForegroundColorColor();
                                if (color != null && color.getHexString().equals(XLS_QUESTION_COLOR) && !gotQuestion) {
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
                                        sections.add(objectMapper.writeValueAsString(section));
                                        section.clear();
                                        gotQuestion = false;
                                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                        body = cell.getRichStringCellValue().getString();
                                        section.put("question", question);
                                        section.put("body", body);
                                        section.putAll(entries);
                                        sections.add(objectMapper.writeValueAsString(section));
                                        section.clear();
                                        gotQuestion = false;

                                    }
                                }
                            }
                            k++;
                        }
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
                while (j < sheet.getLastRowNum()) {
                    XSSFRow row = sheet.getRow(j);

                    int k = 0;
                    if (row != null) {
                        while (k < row.getLastCellNum()) {
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
                                        sections.add(objectMapper.writeValueAsString(section));
                                        section.clear();
                                        gotQuestion = false;
                                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                        body = cell.getRichStringCellValue().getString();
                                        section.put("question", question);
                                        section.put("body", body);
                                        section.putAll(entries);
                                        sections.add(objectMapper.writeValueAsString(section));
                                        section.clear();
                                        gotQuestion = false;

                                    }
                                }
                            }
                            k++;
                        }
                    }

                    j++;
                }
                i++;
            }


        }

        return sections;

    }


    public List<Map<String,Object>> getJsonHM() throws JsonProcessingException {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String,Object> sectionHash = new HashMap<>();
        List<Map<String,Object>> sectionsList = new ArrayList<>();
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
            sectionHash.put("body",body);
            sectionsList.add(sectionHash);
            sectionHash = new HashMap<>();
            body = new StringBuilder();
        }

        return(sectionsList);
    }

}