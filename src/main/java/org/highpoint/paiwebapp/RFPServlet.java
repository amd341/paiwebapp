package org.highpoint.paiwebapp;

/**
 * Created by alex on 6/19/17.
 * Modified by alex and brenden
 * servlet class for ingesting documents to be parsed
 */

import javax.servlet.annotation.MultipartConfig;
import javax.servlet.http.*;
import javax.servlet.ServletException;
import java.io.*;
import java.util.*;


import com.google.gson.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

@MultipartConfig
public class RFPServlet extends HttpServlet {

    /** really basic implementation for GET, allows all origins and simply returns "Request received." */
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws IOException, ServletException {

        PrintWriter out = response.getWriter();
        response.setContentType("text/html; charset=UTF-8");
        response.setHeader("Cache-control", "no-cache, no-store");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "-1");


        response.setHeader("Access-Control-Allow-Origin", "*");
        response.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS,GET");
        response.setHeader("Access-Control-Allow-Headers", "Content-Type");
        response.setHeader("Access-Control-Max-Age", "86400");

        out.println("<h1>Request received</h1>");
        out.close();

    }

    /* When we upload a file and metadata from the frontend, it is sent as multipart/form, which prompts the server
    * to respond with an options request. We want to redirect this options to post. */
    protected void doOptions(HttpServletRequest request, HttpServletResponse response)
            throws IOException, ServletException {
        doPost(request, response);
    }

    /* This takes our file and metadata and does the parsing */
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws IOException, ServletException {


        //get list of allowed origin domains
        List<String> incomingURLs = Arrays.asList(getServletContext().getInitParameter("incomingURLs").trim().split(","));

        //get client's origin domain
        String clientOrigin = request.getHeader("origin");

        PrintWriter out = response.getWriter();
        response.setContentType("application/json; charset=UTF-8");
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

        //gets sections from json body
        Part file = request.getPart("file");
        Part tags = request.getPart("tags");
        Map<String,String> entries = new HashMap<>();

        InputStream tagsInputStream = tags.getInputStream();
        JsonParser jsonParser = new JsonParser();
        JsonObject tagsJsonObject = (JsonObject)jsonParser.parse(new InputStreamReader(tagsInputStream, "UTF-8"));
        tagsInputStream.close();

        boolean useHighlighting = tagsJsonObject.get("useHighlighting").getAsBoolean();

        //gets default tags that aren't the doc body
        for(String key : tagsJsonObject.keySet()) {
            if (!key.equals("companyDoc") && !key.equals("additionalTags") && !key.equals("useHighlighting")) {
                entries.put(key, tagsJsonObject.get(key).getAsString());
            }
        }
        //gets any extra tags they added
        for (JsonElement e : tagsJsonObject.getAsJsonArray("additionalTags")) {
            JsonObject innerPair = e.getAsJsonObject();
            entries.put(innerPair.get("key").getAsString(), innerPair.get("value").getAsString());
        }

        //once we have entries, we have to make sure any new tags get put in as keywords, not text
        SearchClient.mapNewFields("10.20.10.89", 9200, "http", "rfps", "rfp", entries);

        //finding out which kind of office document the RFP is
        List<Map<String,String>> tempstr = null;
        String contentType = file.getContentType();
        if (contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")){
            InputStream filestream = file.getInputStream();
            try {
                ExcelParser parse = new ExcelParser(filestream, entries, true);
                if (useHighlighting) {
                    tempstr = parse.getHighlighted();
                } else {
                    tempstr = parse.getSheets();
                }
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        }
        else if (contentType.equals("application/vnd.ms-excel")){
            InputStream filestream = file.getInputStream();
            try {
                ExcelParser parse = new ExcelParser(filestream, entries, false);
                if (useHighlighting) {
                    tempstr = parse.getHighlighted();

                } else {
                    tempstr = parse.getSheets();
                }
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        }
        else if (contentType.equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document")){
            InputStream filestream = file.getInputStream();
            try {
                DocxParser parse = new DocxParser(filestream, entries);
                if (useHighlighting) {
                    tempstr = parse.getHighlighted();
                } else {
                    tempstr = parse.getSubSections();
                }
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        }
        else{
            System.out.println("ERROR");
        }


        //below here is where, depending on file type, create parser object and then call searchclient methods using
        //returned json string from parse


        JsonObject myObj = new JsonObject();
        myObj.addProperty("success", true);


        JsonObject indexObj = new JsonObject();
        JsonObject emptyObj = new JsonObject();
        indexObj.add("index", emptyObj);

        JsonArray jarray = new JsonArray();

        int counter = 0;
        for(Map<String,String> tempmap: tempstr){
            JsonObject innerObject = new JsonObject();
            for(Map.Entry<String,String> tempentry: tempmap.entrySet()){
                Gson gson = new Gson();
                //replaces weird Word apostrophes with normal ones while adding to json object
                innerObject.addProperty(tempentry.getKey(), tempentry.getValue().toString().replaceAll("[\\u2018\\u2019]", "'")
                        .replaceAll("[\\u201C\\u201D]", "\""));
            }
            jarray.add(indexObj);
            jarray.add(innerObject);
            counter = counter + 1;
        }

        myObj.add("stuff", jarray);
        out.println(myObj);
        out.close();
    }
}
