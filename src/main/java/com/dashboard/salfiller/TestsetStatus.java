package com.dashboard.salfiller;

import lombok.Getter;
import lombok.ToString;

import java.util.LinkedHashMap;
import java.util.List;

@ToString
public class TestsetStatus {

    public final static String PASSED = "Passed";
    public final static String PASSED_WITH_DEFECT = "Passed With Defect";
    public final static String BLOCKED = "Blocked";
    public final static String FAILED = "Failed";
    public final static String NO_RUN = "No Run";
    public final static String NOT_AVAILABLE = "N/A";
    public final static String NOT_COMPLETED = "Not Completed";


    public TestsetStatus(List<Testcase> testcases) {
        numBlocked= testcases.stream().filter(s-> s.getStatus().equals(BLOCKED)).count();
        numFailed= testcases.stream().filter(s-> s.getStatus().equals(FAILED)).count();
        numNA= testcases.stream().filter(s-> s.getStatus().equals(NOT_AVAILABLE)).count();
        numNoRun = testcases.stream().filter(s-> s.getStatus().equals(NO_RUN)).count();
        numNotCompleted= testcases.stream().filter(s-> s.getStatus().equals(NOT_COMPLETED)).count();
        numPassed= testcases.stream().filter(s-> s.getStatus().equals(PASSED)).count();
        numPassedWithDefect= testcases.stream().filter(s-> s.getStatus().equals(PASSED_WITH_DEFECT)).count();
        xlsColumnToStatus.put(2,numBlocked);
        xlsColumnToStatus.put(3,numFailed);
        xlsColumnToStatus.put(4,numNA);
        xlsColumnToStatus.put(5,numNoRun);
        xlsColumnToStatus.put(6,numNotCompleted);
        xlsColumnToStatus.put(7,numPassed);
        xlsColumnToStatus.put(8,numPassedWithDefect);
    }
    @Getter Long numBlocked;
    @Getter Long numPassedWithDefect;
    @Getter Long numPassed;
    @Getter Long numFailed;
    @Getter Long numNA;
    @Getter Long numNotCompleted;
    @Getter Long numNoRun;

    @Getter LinkedHashMap<Integer,Long> xlsColumnToStatus = new LinkedHashMap<>();


}
