package com.dashboard.salfiller;

import lombok.Getter;
import lombok.ToString;
import org.json.JSONObject;

@ToString
public class Testcase {

    @Getter private Integer id;
    @Getter private String status;

    public Testcase(JSONObject testcase) {
        id = Integer.valueOf(testcase.getString("test-id"));
        status = testcase.getString("status");
    }



}
