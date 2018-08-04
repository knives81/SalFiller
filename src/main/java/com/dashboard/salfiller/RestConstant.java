package com.dashboard.salfiller;

import java.util.HashMap;
import java.util.Map;

public final class RestConstant {
	
	public final static String APPLICATION_JSON = "application/json";
	public final static String ACCEPT = "accept";
	public final static String CONTENT = "Content-Type";	
	public final static int OK_STATUS_CODE = 200;
	public final static String ALM_URL = "https://springqcent.saas.hpe.com/qcbin/";
	public final static String COOKIE = "Set-Cookie";

	
	public static final Map<String, String> headerMap;
    static
    {
    	headerMap = new HashMap<String, String>();
    	headerMap.put(ACCEPT, APPLICATION_JSON);
    	headerMap.put(CONTENT, APPLICATION_JSON);	
    }
	
		

}
