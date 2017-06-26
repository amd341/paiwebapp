package org.highpoint.paiwebapp;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.entity.ContentType;
import org.apache.http.nio.entity.NStringEntity;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.elasticsearch.client.Response;
import org.elasticsearch.client.RestClient;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by alex on 6/13/17.
 */
public class ExcelParser {

    private Workbook workbook;
    private Map<String,Object> entries;
    //private static final String FILE_NAME = "/home/alex/documents/excels/enterprise.xlsx";

    public ExcelParser(InputStream input, Map<String,Object> entries, Boolean isXLSX) throws IOException, InvalidFormatException {


        if (isXLSX){
            workbook = new XSSFWorkbook(OPCPackage.open(input));
            this.entries = entries;
        }
        else{
            workbook = new HSSFWorkbook(input);
            this.entries = entries;
        }

    }

    public void printSections(){
        boolean isBlankRow = true;



        Sheet datatypeSheet = workbook.getSheetAt(1);
        Iterator<Row> iterator = datatypeSheet.iterator();

        while (iterator.hasNext()) {

            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();

            while (cellIterator.hasNext()) {

                Cell currentCell = cellIterator.next();

                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                    System.out.print(currentCell.getStringCellValue() + "--");
                    isBlankRow = false;
                }
                else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    System.out.print(currentCell.getNumericCellValue() + "--");
                    isBlankRow = false;
                }
            }
            if (!isBlankRow){
                System.out.println();
                isBlankRow = true;
            }

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

    public String bulkIndex(String hostname, int port, String scheme, String index, String type) throws IOException{
        RestClient restClient = RestClient.builder(new HttpHost(hostname, port, scheme)).build();

        String actionMetaData = String.format("{ \"index\" : { \"_index\" : \"%s\", \"_type\" : \"%s\" } }%n", index, type);

        StringBuilder prepString = new StringBuilder();
        List<String> sectionsList = getJsonStrings();

        for (String s : sectionsList){
            prepString.append(actionMetaData);
            prepString.append(s);
            prepString.append("\n");

        }

        HttpEntity entity = new NStringEntity(prepString.toString(), ContentType.APPLICATION_JSON);

        System.out.println();
        System.out.println(prepString);

        Response response = restClient.performRequest("POST", "/"+index+"/"+type+"/_bulk",
                Collections.emptyMap(), entity);
        restClient.close();

        return response.toString();
    }
}
