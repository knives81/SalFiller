package com.dashboard.salfiller;


import com.mashape.unirest.http.HttpResponse;
import com.mashape.unirest.http.JsonNode;
import com.mashape.unirest.http.Unirest;
import com.mashape.unirest.http.exceptions.UnirestException;
import lombok.Getter;
import lombok.Setter;
import org.springframework.beans.factory.annotation.Autowired;

import java.util.List;
import java.util.stream.Collectors;


public abstract class RestClient {

	protected AlmJsonConverter almJsonConverter = new AlmJsonConverter();

	private static final String LOGIN_ENDPOINT = "api/authentication/sign-in";

	@Getter	@Setter	String domain = "GICT_ITALY_MERCATO";
	@Getter	@Setter	String project = "P59_Change_Deploy_CRMT";
	@Getter	@Setter	String username = "mpinzi";
	@Getter	@Setter	String password = "gTlwd8";
	
	protected String getBaseUrl(String endpoint) {
		return RestConstant.ALM_URL + endpoint +	"/domains/" + domain +	"/projects/" + project + "/";
	}
	
	private String getUrlLogin() {
		return RestConstant.ALM_URL + LOGIN_ENDPOINT;
	}	
	

	protected String getAuthenticationHeader() throws UnirestException {
		String loginUrl = getUrlLogin();
		HttpResponse<JsonNode> postResponse = Unirest.post(loginUrl)
				.basicAuth(username, password)
				.headers(RestConstant.headerMap).asJson();

		if (RestConstant.OK_STATUS_CODE == postResponse.getStatus()) {
			List<String> cookieList = postResponse.getHeaders().get(RestConstant.COOKIE);
			return cookieList.stream().collect(Collectors.joining(";"));
		}
		System.out.println(postResponse.getStatus());
		throw new RuntimeException();
	}

}
