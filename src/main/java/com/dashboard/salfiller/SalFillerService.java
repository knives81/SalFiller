package com.dashboard.salfiller;


import com.mashape.unirest.http.exceptions.UnirestException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.BaseFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

@Component
public class SalFillerService  {
    private static DataFormatter dataFormatter = new DataFormatter();

    public void getTestsetStatus(String domain,String project,String username,String password,String salPath) throws IOException, InvalidFormatException, UnirestException {

        FileInputStream file = new FileInputStream(salPath);
        Workbook workbook = WorkbookFactory.create(file);

        System.out.println("Working Directory = " + System.getProperty("user.dir"));

        fillTestset(workbook);
        fillDefect(workbook);

        //FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        //evaluator.evaluateAll();

        FileOutputStream fileOut = new FileOutputStream(new File(salPath));
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();




    }

    private Workbook fillDefect(Workbook workbook) throws UnirestException {

        RestClientDefect restClientDefect = new RestClientDefect();

        Sheet listaPmoSheet = workbook.getSheet("Lista PMO");
        List<String> AOs = new ArrayList<>();



        for(int rowIndex=4; rowIndex<5; rowIndex++) {
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
        System.out.println(AOs);

        JSONArray releaseCycles = restClientDefect.getReleaseCycles();
        JSONArray releases = restClientDefect.getReleases();
        List<JSONObject> jsonDefects = restClientDefect.getEntityFromAlm(AOs);

        List<Defect> defects = new ArrayList<>();
        for(JSONObject jsonDefect : jsonDefects) {
            Defect defect = new Defect(jsonDefect,releases,releaseCycles);
            defects.add(defect);
        }

        Sheet listaAnomalieAllSheet = workbook.getSheet("ListaAnomalieALL");

        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();


        int rowIndex = 1;
        for(Defect defect : defects) {
            Row row = row = listaAnomalieAllSheet.getRow(rowIndex++);
            for(String column : defect.getXlsColumnToValue().keySet()) {
                Cell cell = row.createCell(CellReference.convertColStringToIndex(column));
                cell.setCellType(CellType.STRING);
                cell.setCellValue(defect.getXlsColumnToValue().get(column));
            }
            /*Cell formulaCell = row.getCell(CellReference.convertColStringToIndex("M"));
            evaluator.evaluateFormulaCell(formulaCell);
            Cell formulaCell2 = row.getCell(CellReference.convertColStringToIndex("N"));
            evaluator.evaluateFormulaCell(formulaCell2);
            Cell formulaCell3 = row.getCell(CellReference.convertColStringToIndex("O"));
            evaluator.evaluateFormulaCell(formulaCell3);*/

        }




        return workbook;

    }

    private void fillTestset(Workbook workbook) throws UnirestException {
        RestClientTestSet restClientTestSet = new RestClientTestSet();

        Sheet testsetIdSheet = workbook.getSheet("testsetid");
        List<Integer> testsetIds = new ArrayList<>();

        for(Row row : testsetIdSheet){
            if(row.getRowNum()>0) {
                Cell testsetIdCell = row.getCell(0);
                String testsetIdCellValue = dataFormatter.formatCellValue(testsetIdCell);
                testsetIds.add(Integer.valueOf(testsetIdCellValue));
            }
        }

        List<JSONArray> testsetsFromAlm = restClientTestSet.getEntityFromAlm(testsetIds);

        List<List<Testcase>> testcasePerTestSet = new ArrayList<>();
        for(JSONArray testsetFromAlm : testsetsFromAlm) {
            List<Testcase> testcases = new ArrayList<>();
            for(int i=0; i<testsetFromAlm.length(); i++) {
                Testcase testcase = new Testcase(testsetFromAlm.getJSONObject(i));
                testcases.add(testcase);
                System.out.println(testcase);
            }
            testcasePerTestSet.add(testcases);
        }

        List<TestsetStatus> testsetsStatus = new ArrayList<>();
        testcasePerTestSet.stream().forEach(t -> testsetsStatus.add(new TestsetStatus(t)));

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
}
