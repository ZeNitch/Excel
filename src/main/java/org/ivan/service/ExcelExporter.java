/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.service;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.ivan.model.ValueObject;
import org.springframework.stereotype.Service;

/**
 *
 * @author ivan
 */
@Service
public class ExcelExporter {

    private static final String NEW_LINE_SEPARATOR = "\n";

    public <T extends ValueObject> File createExcelExport(List<T> objects, String fileName) throws IOException {

        File xlsFile = new File(fileName);
        FileWriter writer = new FileWriter(xlsFile);

        BufferedWriter bufferedWriter = new BufferedWriter(writer);
        List<String> headers = objects.get(0).getHeaders();
        for (String header : headers) {
            bufferedWriter.write(header);
        }
        bufferedWriter.newLine();
        List<String> values = new ArrayList<>();
        for (T obj : objects) {
            //bufferedWriter.write(obj.getValues());
            values = obj.getValues();
            for (String value : values) {
                bufferedWriter.write(value);
            }
            bufferedWriter.newLine();
        }

        bufferedWriter.close();

        return xlsFile;
    }

    public <T extends ValueObject> File createExcelExportDeprecated(List<T> objects, String fileName) throws IOException {

        Workbook workbook;
        if (FilenameUtils.getExtension(fileName).equals("xls")) {
            workbook = new HSSFWorkbook();
        } else if (FilenameUtils.getExtension(fileName).equals("xlsx")) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new HSSFWorkbook();
            fileName += ".xls";
        }

        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        int cellNum = 0;

        List<String> headers = objects.get(0).getHeaders();
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(cellNum++);
            cell.setCellValue(headers.get(i));
        }
        int rowNum = 1;
        List<String> values = new ArrayList<>();// = objects.get(0).getValues().split("\t");
        for (T obj : objects) {

            values = obj.getValues();
            cellNum = 0;
            row = sheet.createRow(rowNum++);
            for (int i = 0; i < values.size(); i++) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(values.get(i));
            }

        }

        return excelFile(workbook, fileName);
    }

    /**
     *
     * @param classDefinition
     * @return
     */
    public List<String> getFieldNames(Class classDefinition) {

        List<String> fieldNames = new ArrayList<>();

        for (Field field : classDefinition.getDeclaredFields()) {
            fieldNames.add(field.getName());
        }

        return fieldNames;
    }

    public <E extends ValueObject> Map<String, List<List<String>>> getSheetListsByFieldName(List<E> objects, String fieldName) {

        Map<String, List<List<String>>> sheetList = new HashMap<>();
        //Map<List<List String>>> sheetList = new HashMap<>();
        for (E object : objects) {
            Map<String, String> objectValues = getValuesWithHeaders(object);

            if (!sheetList.containsKey(objectValues.get(fieldName))) {
                sheetList.put(objectValues.get(fieldName), new ArrayList<>());
            }
            sheetList.get(objectValues.get(fieldName)).add(new ArrayList<>(objectValues.values()));
        }
        return sheetList;
    }

    public <T extends ValueObject> Map<String, String> getValuesWithHeaders(T object) {

        Map<String, String> valuesWithHeaders = new LinkedHashMap<>();

        List<String> headers = object.getHeaders();
        List<String> values = object.getValues();

        for (int i = 0; i < headers.size(); i++) {
            valuesWithHeaders.put(headers.get(i), values.get(i));
        }

        return valuesWithHeaders;
    }
    // Pattern for generic type usage
    //    public <T extends Object> T getObject() {
    //        return null;
    //    }

    public <T extends Object, E extends Object> File createExcelFromSheetMap(List<ValueObject> objects, String fieldName, String fileName) {

        Workbook workbook;
        if (FilenameUtils.getExtension(fileName).equals("xls")) {
            workbook = new HSSFWorkbook();
        } else if (FilenameUtils.getExtension(fileName).equals("xlsx")) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new HSSFWorkbook();
            fileName += ".xls";
        }
        Map<String, List<List<String>>> sheetLists = getSheetListsByFieldName(objects, fieldName);
        sheetLists.keySet().forEach((mapKey) -> {
            Sheet sheet = workbook.createSheet(mapKey.toString());
            putHeaders(objects.get(0).getHeaders(), sheet, workbook);
            putCellValues(sheet, sheetLists.get(mapKey));
        });

        return excelFile(workbook, fileName);
    }

    private void putHeaders(List<String> headers, Sheet sheet, Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);

        Row row = sheet.createRow(0);
        int cellNum = 0;
        for (String header : headers) {
            Cell cell = row.createCell(cellNum++);
            cell.setCellValue(header);
            cell.setCellStyle(style);
        }
    }

    private void putCellValues(Sheet sheet, List<List< String>> cellValues) {
        int rowNum = 1;
        int cellNum;
        Row row;
        for (List<String> objectValues : cellValues) {
            cellNum = 0;
            row = sheet.createRow(rowNum++);
            for (String parameter : objectValues) {

                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(parameter);
            }
        }
    }

    private File excelFile(Workbook workbook, String fileName) {
        File file = new File(fileName);
        try {
            try (FileOutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
            }

        } catch (FileNotFoundException e) {
        } catch (IOException e) {
        }

        return file;
    }
}
