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
import java.util.stream.Collectors;

@Component("defect")
public class RestClientDefect extends RestClient {

	private static final String DEFECTS = "defects";
	private static final String RELEASE_CYCLES = "release-cycles";
	private static final String RELEASES = "releases";

	private static final Logger log = LoggerFactory.getLogger(RestClientDefect.class);
	
	protected String getUrlDefect() {
		return getBaseUrl("api") + DEFECTS;
	}

	public List<JSONObject> getEntityFromAlm(List<String> AOs) throws UnirestException {
		String authenticationHeader = this.getAuthenticationHeader();


		List<JSONObject> defects = new ArrayList<>();
		for (String AO : AOs) {

			JSONArray defectFromAlm = getDefect(authenticationHeader, AO);
			for (int i = 0; i < defectFromAlm.length(); i++) {
				JSONObject defect = defectFromAlm.getJSONObject(i);
				defects.add(defect);
			}

		}
		
		return defects;
	}


	private JSONArray getDefect(String authenticationHeader, String AO) throws UnirestException {
		String urlDefect = getUrlDefect();
		HttpResponse<JsonNode> getResponseDefect = Unirest.get(urlDefect)
				.queryString("query", "\\\"user-template-02='"+AO+"'\\\"" )
				.queryString("limit", 300)
				.headers(RestConstant.headerMap)
				.header("Cookie", authenticationHeader).asJson();
		if (RestConstant.OK_STATUS_CODE == getResponseDefect.getStatus()) {

			Integer numberOfEntities = getResponseDefect.getBody().getObject().getJSONArray("data").length();
			log.info(urlDefect + "query:" + AO + " Entities:"+ numberOfEntities);

			return getResponseDefect.getBody().getObject().getJSONArray("data");
		} else {
			throw new RuntimeException("Error get defect");
		}
	}

	public JSONArray getReleases() throws UnirestException {
		return getEntity(RELEASES);
	}
	public JSONArray getReleaseCycles() throws UnirestException {
		return getEntity(RELEASE_CYCLES);
	}

	private JSONArray getEntity(String entityName) throws UnirestException {
		String entityUrl = getEntityUrl(entityName);
		String authenticationHeader = getAuthenticationHeader();

		HttpResponse<JsonNode> getEntityUrl = Unirest.get(entityUrl)
				.queryString("limit", 150)
				.headers(RestConstant.headerMap)
				.header("Cookie", authenticationHeader).asJson();
		if (RestConstant.OK_STATUS_CODE == getEntityUrl.getStatus()) {
			return almJsonConverter.convertEntities(getEntityUrl.getBody().getObject());
		} else {
			throw new RuntimeException();
		}
	}
	private String getEntityUrl(String entityName) {
		return getBaseUrl("rest") + entityName;
	}
}
