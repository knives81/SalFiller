package com.dashboard.salfiller;


import com.mashape.unirest.http.exceptions.UnirestException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Component
public class SalFillerService  {
    private static DataFormatter dataFormatter = new DataFormatter();

    public void getTestsetStatus(String domain,String project,String username,String password,String salPath) throws IOException, InvalidFormatException, UnirestException {

        Workbook workbook = WorkbookFactory.create(new FileInputStream(salPath));

        fillTestset(workbook);
        fillDefect(workbook);

        saveModifiedSal(salPath, workbook);
    }

    private void saveModifiedSal(String salPath, Workbook workbook) throws IOException {
        FileOutputStream fileOut = new FileOutputStream(new File(salPath));
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    private Workbook fillDefect(Workbook workbook) throws UnirestException {

        RestClientDefect restClientDefect = new RestClientDefect();

        List<String> AOs = findAOsInListaPMO(workbook);

        List<Defect> defects = buildDefects(restClientDefect, AOs);

        putDefectsInListaAnomalieALL(workbook, defects);

        return workbook;

    }

    private void putDefectsInListaAnomalieALL(Workbook workbook, List<Defect> defects) {
        Sheet listaAnomalieAllSheet = workbook.getSheet("ListaAnomalieALL");
        int rowIndex = 1;
        for(Defect defect : defects) {
            Row row = row = listaAnomalieAllSheet.getRow(rowIndex++);
            for(String column : defect.getXlsColumnToValue().keySet()) {
                Cell cell = row.createCell(CellReference.convertColStringToIndex(column));
                cell.setCellType(CellType.STRING);
                cell.setCellValue(defect.getXlsColumnToValue().get(column));
            }
        }
        System.out.println(defects.size() +" defects found");
    }

    private List<Defect> buildDefects(RestClientDefect restClientDefect, List<String> AOs) throws UnirestException {
        JSONArray releaseCycles = restClientDefect.getReleaseCycles();
        JSONArray releases = restClientDefect.getReleases();
        List<JSONObject> jsonDefects = restClientDefect.getEntityFromAlm(AOs);

        List<Defect> defects = new ArrayList<>();
        for(JSONObject jsonDefect : jsonDefects) {
            Defect defect = new Defect(jsonDefect,releases,releaseCycles);
            defects.add(defect);
        }
        return defects;
    }

    private List<String> findAOsInListaPMO(Workbook workbook) {
        Sheet listaPmoSheet = workbook.getSheet("Lista PMO");
        List<String> AOs = new ArrayList<>();

        for(int rowIndex=4; rowIndex<30; rowIndex++) {
            Row row = listaPmoSheet.getRow(rowIndex);
            Cell AOCell = row.getCell(CellReference.convertColStringToIndex("B"));
            String AO = "";
            if(AOCell==null) {
                break;
            } else {
                AO = dataFormatter.formatCellValue(AOCell);
            }
            if(StringUtils.isNotEmpty(AO)) {
                AOs.add(AO);
            }
        }
        System.out.println("AOs in Lista PMO");
        System.out.println(AOs);
        return AOs;
    }

    private void fillTestset(Workbook workbook) throws UnirestException {


        Sheet testsetIdSheet = workbook.getSheet("testsetid");
        List<Integer> testsetIds = getTestsetIdsFromXls(testsetIdSheet);

        List<JSONArray> testsetsFromAlm = getTestsetsFromAlm(testsetIds);

        List<TestsetStatus> testsetsStatus = getTestsetStatus(testsetsFromAlm);

        putTestsetStatusInXls(testsetIdSheet, testsetsStatus);
    }

    private void putTestsetStatusInXls(Sheet testsetIdSheet, List<TestsetStatus> testsetsStatus) {
        int rowIndex = 1;
        for(TestsetStatus testsetStatus : testsetsStatus) {
            Row row = row = testsetIdSheet.getRow(rowIndex++);
            for(Integer columnId : testsetStatus.getXlsColumnToStatus().keySet()) {
                Cell cell = row.createCell(columnId);
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(testsetStatus.getXlsColumnToStatus().get(columnId));
            }
        }
    }

    private List<TestsetStatus> getTestsetStatus(List<JSONArray> testsetsFromAlm) {
        List<List<Testcase>> testcasePerTestSet = new ArrayList<>();
        for(JSONArray testsetFromAlm : testsetsFromAlm) {
            List<Testcase> testcases = new ArrayList<>();
            for(int i=0; i<testsetFromAlm.length(); i++) {
                Testcase testcase = new Testcase(testsetFromAlm.getJSONObject(i));
                testcases.add(testcase);
            }
            testcasePerTestSet.add(testcases);
        }
        List<TestsetStatus> testsetsStatus = new ArrayList<>();
        testcasePerTestSet.stream().forEach(t -> testsetsStatus.add(new TestsetStatus(t)));
        return testsetsStatus;
    }

    private List<JSONArray> getTestsetsFromAlm(List<Integer> testsetIds) throws UnirestException {
        RestClientTestSet restClientTestSet = new RestClientTestSet();
        return restClientTestSet.getEntityFromAlm(testsetIds);
    }

    private List<Integer> getTestsetIdsFromXls(Sheet testsetIdSheet) {
        List<Integer> testsetIds = new ArrayList<>();
        for(Row row : testsetIdSheet){
            if(row.getRowNum()>0) {
                Cell testsetIdCell = row.getCell(0);
                String testsetIdCellValue = dataFormatter.formatCellValue(testsetIdCell);
                testsetIds.add(Integer.valueOf(testsetIdCellValue));
            }
        }
        return testsetIds;
    }
}
