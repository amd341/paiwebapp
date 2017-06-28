package org.highpoint.paiwebapp;

/**
 * Created by alex on 6/19/17.
 */

import javax.servlet.annotation.MultipartConfig;
import javax.servlet.http.*;
import javax.servlet.ServletException;
import java.io.*;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

@MultipartConfig
public class HelloServlet extends HttpServlet {

    private enum Choice{
        DOCX, XLS, XLSX, ERROR
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws IOException, ServletException {
        doPost(request, response);
    }

    protected void doOptions(HttpServletRequest request, HttpServletResponse response)
            throws IOException, ServletException {
        doPost(request, response);
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws IOException, ServletException {


        //get list of allowed origin domains
        List<String> incomingURLs = Arrays.asList(getServletContext().getInitParameter("incomingURLs").trim().split(","));

        //get client's origin domain
        String clientOrigin = request.getHeader("origin");
        System.out.println(clientOrigin);

        PrintWriter out = response.getWriter();
        response.setContentType("application/json");
        response.setHeader("Cache-control", "no-cache, no-store");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "-1");

        //checking if client's origin is allowed
        int originCheckIndex = incomingURLs.indexOf(clientOrigin);
        if (originCheckIndex != -1){
            response.setHeader("Access-Control-Allow-Origin", clientOrigin);
            response.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS,GET");
            response.setHeader("Access-Control-Allow-Headers", "Content-Type");
            response.setHeader("Access-Control-Max-Age", "86400");

        }
        //System.out.println(IOUtils.toString(request.getReader()));

        
        //String type = request.getParameter("type");
        //String age = request.getParameter("age");
        Part file = request.getPart("file");
        Part tags = request.getPart("tags");
        Map<String,Object> entries = new HashMap<>();

        InputStream tagsInputStream = tags.getInputStream();
        JsonParser jsonParser = new JsonParser();
        JsonObject tagsJsonObject = (JsonObject)jsonParser.parse(new InputStreamReader(tagsInputStream, "UTF-8"));
        tagsInputStream.close();

        for(String key : tagsJsonObject.keySet()) {
            if (!key.equals("companyDoc") && !key.equals("additionalTags")) {
                entries.put(key, tagsJsonObject.get(key).getAsString());
            }
        }
        for (JsonElement e : tagsJsonObject.getAsJsonArray("additionalTags")) {
            JsonObject innerPair = e.getAsJsonObject();
            entries.put(innerPair.get("key").getAsString(), innerPair.get("value").getAsString());
        }


        //finding out which kind of office document the RFP is
        //Choice choice;
        List<Map<String,Object>> tempstr = null;
        String contentType = file.getContentType();
        if (contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")){
            System.out.println("XLSX");
            //choice = Choice.XLSX;
            InputStream filestream = file.getInputStream();
            try {
                ExcelParser parse = new ExcelParser(filestream, entries, true);
                tempstr = parse.getJsonHM();
                //System.out.println(tempstr);
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        }
        else if (contentType.equals("application/vnd.ms-excel")){
            System.out.println("XLS");
            InputStream filestream = file.getInputStream();
            try {
                ExcelParser parse = new ExcelParser(filestream, entries, false);
                tempstr = parse.getJsonHM();
                //System.out.println(tempstr);
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        }
        else if (contentType.equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document")){
            System.out.println("DOCX");
            //choice = Choice.DOCX;
            InputStream filestream = file.getInputStream();
            try {
                DocxParser parse = new DocxParser(filestream, entries);
                tempstr = parse.getSections();
                //System.out.println(tempstr);
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        }
        else{
            System.out.println("ERROR");
            //choice = Choice.ERROR;
        }


        //below here is where, depending on file type, create parser object and then call searchclient methods using
        //returned json string from parse


        JsonObject myObj = new JsonObject();
        myObj.addProperty("success", true);



        JsonArray jarray = new JsonArray();

        int counter = 0;
        for(Map<String,Object> tempmap: tempstr){
            JsonObject innerObject = new JsonObject();
            for(Map.Entry<String,Object> tempentry: tempmap.entrySet()){
                //RIGHT HERE BUDDY
                innerObject.addProperty(tempentry.getKey(), tempentry.getValue().toString());
                System.out.println(innerObject);
            }
            //System.out.println(jarray);
            jarray.add(innerObject);
            System.out.println(jarray);
            counter = counter + 1;
        }

        myObj.add("stuff", jarray);

        //myObj.addProperty("stuff", tempstr.toString());
        System.out.println(myObj);
        out.println(myObj);


        out.close();
    }
}
