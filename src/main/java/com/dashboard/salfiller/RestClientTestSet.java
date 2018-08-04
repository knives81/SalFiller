package com.dashboard.salfiller;

import com.mashape.unirest.http.HttpResponse;
import com.mashape.unirest.http.JsonNode;
import com.mashape.unirest.http.Unirest;
import com.mashape.unirest.http.exceptions.UnirestException;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component("testset")
public class RestClientTestSet extends RestClient {
	
	
	private static final String TEST_INSTANCES = "test-instances";




	private static final Logger log = LoggerFactory.getLogger(RestClientTestSet.class);
	
	protected String getUrlTestInstances() {
		return getBaseUrl("rest") + TEST_INSTANCES;
	}

	public List<JSONArray> getEntityFromAlm(List<Integer> testsetIds) throws UnirestException {
		String authenticationHeader = this.getAuthenticationHeader();

		List<JSONArray> testsetsFromAlm = new ArrayList<>();
		for (Integer testsetId : testsetIds) {
			JSONArray testsetFromAlm = getTestInstance(authenticationHeader, testsetId);
			testsetsFromAlm.add(testsetFromAlm);
		}
		return testsetsFromAlm;
	}

	

	private JSONArray getTestInstance(String authenticationHeader, Integer testsetId)
			throws UnirestException {
		String urlTestInstances = getUrlTestInstances();

		HttpResponse<JsonNode> getResponseTestInstances = Unirest.get(urlTestInstances)
				.queryString("query", "{cycle-id[" + testsetId + "]}")
				.headers(RestConstant.headerMap)
				.header("Cookie", authenticationHeader).asJson();
		if (RestConstant.OK_STATUS_CODE == getResponseTestInstances.getStatus()) {
			Integer numberOfEntities = getResponseTestInstances.getBody().getObject().getJSONArray("entities").length();
			log.info(urlTestInstances + " {cycle-id[" + testsetId + "]} Entities:"+ numberOfEntities);
			JSONArray arrayOfTest = almJsonConverter.convertEntities(getResponseTestInstances.getBody().getObject());
			return  arrayOfTest;

		} else {
			throw new RuntimeException();
		}
	}
	
}
