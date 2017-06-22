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
import org.apache.commons.io.IOUtils;

public class HelloServlet extends HttpServlet {

    private enum choice{
        DOCX, XLS, XLSX
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
        response.setContentType("text/html");
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

        System.out.println(IOUtils.toString(request.getReader()));
        //String type = request.getParameter("type");
        //String age = request.getParameter("age");

        JsonObject myObj = new JsonObject();
        myObj.addProperty("success", true);
        myObj.addProperty("secondproperty", "george");
        out.println(myObj.toString());


        out.close();
    }
}
