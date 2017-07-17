package org.highpoint.paiwebapp;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.http.HttpEntity;
import org.apache.http.HttpHost;
import org.apache.http.entity.ContentType;
import org.apache.http.nio.entity.NStringEntity;
import org.apache.http.util.EntityUtils;
import org.elasticsearch.client.Response;
import org.elasticsearch.client.RestClient;

import java.io.IOException;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Created by brenden on 6/20/17.
 * SearchClient class with static methods for bulk indexing
 * and mapping
 */
public class SearchClient {

    /**
     * bulk indexes sections from word document
     * @param hostname hostname of elasticsearch instance
     * @param port port for accessing
     * @param scheme scheme of access
     * @param index name of index to add data to
     * @param type name of type to add data to
     * @param bulkData list of json strings to be indexed
     * @return String response from elasticsearch
     * @throws IOException if hostname is incorrect
     */
    public static String bulkIndex(String hostname, int port, String scheme, String index, String type, List<String> bulkData) throws IOException {
        RestClient restClient = RestClient.builder(
                new HttpHost(hostname, port, scheme)).build();

        String actionMetaData = String.format("{ \"index\" : { \"_index\" : \"%s\", \"_type\" : \"%s\" } }%n", index, type);

        StringBuilder bulkRequestBody = new StringBuilder();
        for (String bulkItem : bulkData) {
            bulkRequestBody.append(actionMetaData);
            bulkRequestBody.append(bulkItem);
            bulkRequestBody.append("\n");
        }
        HttpEntity entity = new NStringEntity(bulkRequestBody.toString(), ContentType.APPLICATION_JSON);

        Response response = restClient.performRequest("POST", "/"+index+"/"+type+"/_bulk",
                Collections.emptyMap(), entity);
        restClient.close();
        return response.toString();
    }

    /**
     * adds any new keys as keywords to elasticsearch. this is assuming there is already a default
     * mapping which should have basic, essential fields like body or date in whichever datatype you prefer
     * @param hostname hostname of elasticsearch instance
     * @param port port for accessing
     * @param scheme scheme of access
     * @param index name of index to add data to
     * @param type name of type to add data to
     * @param entries fields to be extracted and added
     * @return String response from server
     * @throws IOException if connection or request is bad
     */
    public static String mapNewFields(String hostname, int port, String scheme, String index, String type, Map<String,String> entries) throws IOException {
        RestClient restClient = RestClient.builder(
                new HttpHost(hostname, port, scheme)).build();

        Response response = restClient.performRequest("GET", "/"+index+"/"+type+"/_mapping");
        ObjectMapper objectMapper = new ObjectMapper();

        JsonNode root = objectMapper.readTree(EntityUtils.toString(response.getEntity()));

        for (String key : entries.keySet())
        {
            if(root.findPath(key).isMissingNode())
            {
                String newMapping = String.format("{ \"properties\": { \"%s\": { \"type\": \"keyword\" } } } ", key);
                HttpEntity entity = new NStringEntity(newMapping, ContentType.APPLICATION_JSON);
                response = restClient.performRequest("PUT", "/"+index+"/"+type+"/_mapping", Collections.emptyMap(), entity);
            }
        }
        restClient.close();
        return EntityUtils.toString(response.getEntity());

    }

}
