package com.dashboard.salfiller;

import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Component;

@Component
public class AlmJsonConverter {
	
	private static final String  ENTITIES = "entities";
	private static final String  FIELDS = "Fields";
	private static final String  NAME = "Name";
	private static final String  VALUES = "values";
	private static final String  VALUE = "value";
	private static final String  EMPTY = "";
	
	
	public JSONArray convertEntities(JSONObject jsonRestResponse) {		
			
			JSONArray jsonEntities = jsonRestResponse.getJSONArray(ENTITIES);
			JSONArray jsonNewEntities = new JSONArray();
			for(int i=0; i<jsonEntities.length(); i++) {
				JSONObject jsonEntity = jsonEntities.getJSONObject(i);
				JSONObject jsonNewEntity = createEntity(jsonEntity);
				jsonNewEntities.put(jsonNewEntity);
			}
			return jsonNewEntities;
	}
	
	public JSONObject convertEntity(JSONObject jsonRestResponse) {
		return createEntity(jsonRestResponse);
	}

	private JSONObject createEntity(JSONObject jsonEntity) {		
		JSONArray jsonFields = jsonEntity.getJSONArray(FIELDS);		
		JSONObject jsonNewEntity = new JSONObject();
		for(int j=0; j<jsonFields.length(); j++) {					
			JSONObject jsonField = jsonFields.getJSONObject(j);
			
			String name = extractName(jsonField);			
			String value = extractValue(jsonField);
			
			jsonNewEntity.put(name,value);
		}
		return jsonNewEntity;
	}

	private String extractValue(JSONObject jsonField) {
		JSONArray jsonValues = jsonField.getJSONArray(VALUES);
		return extractValue(jsonValues);
	}

	private String extractName(JSONObject jsonField) {
		String name = jsonField.getString(NAME);
		return name;
	}

	private String extractValue(JSONArray jsonValues) {
		String value = EMPTY;
		if(jsonValues.length()==1) {
			if(jsonValues.getJSONObject(0).length()==0){
				value = EMPTY;
				
			}
			else if(jsonValues.getJSONObject(0).length()==1) {
				value = jsonValues.getJSONObject(0).getString(VALUE);
				
			}
			else if(jsonValues.getJSONObject(0).length()==2) {
				value = jsonValues.getJSONObject(0).toString();
				
			}
			else{
				assert(false);
			}
		}
		else if (jsonValues.length()==0){
			value = EMPTY;
			
		}
		else {
			assert(false);
		}
		return value;
	}

}
