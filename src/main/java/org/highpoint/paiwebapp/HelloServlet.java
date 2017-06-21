package org.highpoint.paiwebapp;

/**
 * Created by alex on 6/19/17.
 */

import javax.servlet.http.*;
import javax.servlet.ServletException;
import java.io.PrintWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import com.google.gson.JsonObject;

public class HelloServlet extends HttpServlet {

    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws IOException {
        doPost(request, response);
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws IOException {

        //get list of allowed origin domains
        List<String> incomingURLs = Arrays.asList(getServletContext().getInitParameter("incomingURLs").trim().split(","));

        //get client's origin domain
        String clientOrigin = request.getHeader("origin");

        PrintWriter out = response.getWriter();
        response.setContentType("text/html");
        response.setHeader("Cache-control", "no-cache, no-store");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "-1");

        //checking if client's origin is allowed
        int originCheckIndex = incomingURLs.indexOf(clientOrigin);
        if (originCheckIndex != -1){
            response.setHeader("Access-Control-Allow-Origin", clientOrigin);
            response.setHeader("Access-Control-Allow-Methods", "POST");
            response.setHeader("Access-Control-Allow-Headers", "Content-Type");
            response.setHeader("Access-Control-Max-Age", "86400");

        }
        JsonObject myObj = new JsonObject();
        myObj.addProperty("success", true);
        myObj.addProperty("secondproperty", "george");
        out.println(myObj.toString());


        out.close();
    }
}
