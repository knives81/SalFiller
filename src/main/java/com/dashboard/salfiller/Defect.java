package com.dashboard.salfiller;


import lombok.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.LinkedHashMap;
import java.util.Map;


@ToString
public class Defect {

    public final static String DEFECT_ID = "id";
    public final static String PMO_ID = "user-template-02";
    public final static String SUMMARY = "name";
    public final static String STATUS = "status";
    public final static String ASSIGNED_TO = "owner";
    public final static String DETECTED_IN_CYCLE = "detected-in-rcyc";
    public final static String DETECTED_IN_RELEASE = "detected-in-rel";
    public final static String PRIORITY = "priority";
    public final static String SEVERITY = "severity";
    public final static String SISTEMI = "user-template-01";
    public final static String TIPOLOGIA_DI_DEFECT = "user-01";
    public final static String DETECTED_ON_DATE = "creation-time";

    private static final Map<String, Defect.FieldType> keyTotype;
    static
    {
        keyTotype = new LinkedHashMap<String, Defect.FieldType>();
        keyTotype.put(DEFECT_ID, FieldType.INT);
        keyTotype.put(PMO_ID, FieldType.STRING);
        keyTotype.put(SUMMARY, FieldType.STRING);
        keyTotype.put(STATUS, FieldType.STRING);
        keyTotype.put(ASSIGNED_TO, FieldType.STRING);
        keyTotype.put(DETECTED_IN_CYCLE, FieldType.RELEASECYCLE);
        keyTotype.put(DETECTED_IN_RELEASE, FieldType.RELEASE);
        keyTotype.put(PRIORITY, FieldType.STRING);
        keyTotype.put(SEVERITY, FieldType.STRING);
        keyTotype.put(SISTEMI, FieldType.STRING);
        keyTotype.put(TIPOLOGIA_DI_DEFECT, FieldType.STRING);
        keyTotype.put(DETECTED_ON_DATE, FieldType.STRING);
    }
    private enum FieldType {
        INT("int"),
        STRING("string"),
        RELEASE("release"),
        RELEASECYCLE("releaseCycle");
        private String value;
        FieldType(String value) {
            this.value = value;
        }
        public String value() {
            return value;
        }
    }


    public Defect(JSONObject defect, JSONArray releases, JSONArray releaseCycles) {
        defectID = getValue(DEFECT_ID,defect,releases, releaseCycles);
        pmoId = getValue(PMO_ID,defect,releases, releaseCycles);
        summary = getValue(SUMMARY,defect,releases, releaseCycles);
        status = getValue(STATUS,defect,releases, releaseCycles);
        assignedTo = getValue(ASSIGNED_TO,defect,releases, releaseCycles);
        detectedInCycle = getValue(DETECTED_IN_CYCLE,defect,releases, releaseCycles);
        detectedInRelease = getValue(DETECTED_IN_RELEASE,defect,releases, releaseCycles);
        priority = getValue(PRIORITY,defect,releases, releaseCycles);
        severity = getValue(SEVERITY,defect,releases, releaseCycles);
        sistemi = getValue(SISTEMI,defect,releases, releaseCycles);
        tipologiaDiDefect = getValue(TIPOLOGIA_DI_DEFECT,defect,releases, releaseCycles);
        detectedOnDate = getValue(DETECTED_ON_DATE,defect,releases, releaseCycles);

        xlsColumnToValue.put("A",defectID);
        xlsColumnToValue.put("B",pmoId);
        xlsColumnToValue.put("C",summary);
        xlsColumnToValue.put("D",status);
        xlsColumnToValue.put("E",assignedTo);
        xlsColumnToValue.put("F",detectedInCycle);
        xlsColumnToValue.put("G",detectedInRelease);
        xlsColumnToValue.put("H",priority);
        xlsColumnToValue.put("I",severity);
        xlsColumnToValue.put("J",sistemi);
        xlsColumnToValue.put("K",tipologiaDiDefect);
        xlsColumnToValue.put("L",detectedOnDate);
        
    }
    @Getter String defectID;
    @Getter String pmoId;
    @Getter String summary;
    @Getter String status;
    @Getter String assignedTo;
    @Getter String detectedInCycle;
    @Getter String detectedInRelease;
    @Getter String priority;
    @Getter String severity;
    @Getter String sistemi;
    @Getter String tipologiaDiDefect;
    @Getter String detectedOnDate;

    @Getter LinkedHashMap<String,String> xlsColumnToValue = new LinkedHashMap<>();

    private String getValue(String key, JSONObject defect, JSONArray releases, JSONArray releaseCycles) {
        if(keyTotype.get(key).equals(FieldType.INT)) {
            return String.valueOf(defect.getInt(key));
        }
        else if(keyTotype.get(key).equals(FieldType.STRING)) {
            try {
                return defect.getString(key);
            } catch(JSONException e) {
                return "";
            }
        }
        else if(keyTotype.get(key).equals(FieldType.RELEASE)) {
            return this.findValue(releases, defect.getJSONObject(DETECTED_IN_RELEASE).getInt("id"));
        }
        else if(keyTotype.get(key).equals(FieldType.RELEASECYCLE)) {
            return this.findValue(releaseCycles, defect.getJSONObject(DETECTED_IN_CYCLE).getInt("id"));
        }
        throw new RuntimeException("No type found");
    }

    private String findValue(JSONArray jsonarray, int value) {
        for(int i=0; i<jsonarray.length(); i++) {
            JSONObject json = jsonarray.getJSONObject(i);
            if(json.getInt("id")==value) {
                return json.getString("name");
            }
        }
        throw new RuntimeException("No name found");

    }
}
