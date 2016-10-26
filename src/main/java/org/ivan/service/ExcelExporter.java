/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.service;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.ivan.model.Person;
import org.springframework.stereotype.Service;

/**
 *
 * @author ivan
 */
@Service
public class ExcelExporter {

    private static final String NEW_LINE_SEPARATOR = "\n";

    public File createExcelExport(List<Person> objects, List<String> headers, String fileName) throws IOException {

        File xlsFile = new File(fileName);
        FileWriter writer = new FileWriter(xlsFile);

        BufferedWriter bufferedWriter = new BufferedWriter(writer);
        for (int i = 0; i < headers.size(); i++) {
            String cup = headers.get(i);

            if (i < headers.size() - 1) {
                cup += "\t";
            }
            bufferedWriter.write(cup);
        }
        bufferedWriter.newLine();
        for (Person obj : objects) {
            bufferedWriter.write(obj.getPersonId() + "\t");
            bufferedWriter.write(obj.getPersonName() + "\t");
            bufferedWriter.write(obj.getGender() + "\t");
            bufferedWriter.write(obj.getAddress().getAddressId() + "\t");
            bufferedWriter.write(obj.getAddress().getStreet() + "\t");
            bufferedWriter.write(obj.getAddress().getCity() + "\t");
            bufferedWriter.write(obj.getAddress().getCountry() + "\n");
        }

        bufferedWriter.close();

        return xlsFile;
    }

    public List<String> getFieldNames(Class classDefinition) {

        List<String> fieldNames = new ArrayList<>();

        for (Field field : classDefinition.getDeclaredFields()) {
            fieldNames.add(field.getName());
        }

        return fieldNames;
    }

    public <E extends Object> Map<String, List<Map<String, String>>> getSheetListsByFieldName(List<E> objects, String fieldName) {

        Map<String, List<Map<String, String>>> sheetList = new HashMap<>();
        for (E object : objects) {
            Map<String, String> objectValues = getValuesWithHeaders(object);

            if (!sheetList.containsKey(objectValues.get(fieldName))) {
                sheetList.put(objectValues.get(fieldName), new ArrayList<>());
            }
            sheetList.get(objectValues.get(fieldName)).add(objectValues);
        }
        return sheetList;
    }

    public <T extends Object> Map<String, String> getValuesWithHeaders(T obj) {

        List<String> fieldNames = getFieldNames(obj.getClass());
        PropertyDescriptor propertyDescriptor;
        Method currentGetMethod;
        Map<String, String> valuesWithHeaders = new HashMap<>();
        for (String fieldName : fieldNames) {
            try {
                propertyDescriptor = new PropertyDescriptor(fieldName, obj.getClass());
                currentGetMethod = propertyDescriptor.getReadMethod();

                if (currentGetMethod.invoke(obj).toString().contains("=")) {
                    valuesWithHeaders.putAll(getValuesWithHeaders(currentGetMethod.invoke(obj)));
                } else {
                    valuesWithHeaders.put(fieldName, currentGetMethod.invoke(obj).toString());
                }

            } catch (IntrospectionException | IllegalAccessException | IllegalArgumentException | InvocationTargetException ex) {
                Logger.getLogger(ExcelExporter.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return valuesWithHeaders;
    }
// Pattern for generic type usage
//    public <T extends Object> T getObject() {
//        return null;
//    }

    public <T extends Object, E extends Object> File createExcelFromSheetMap(Map<T, List<Map<String, String>>> sheetLists, String fileName) {

        Workbook workbook;
        if (FilenameUtils.getExtension(fileName).equals("xls")) {
            workbook = new HSSFWorkbook();
        } else if (FilenameUtils.getExtension(fileName).equals("xlsx")) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new HSSFWorkbook();
            fileName += ".xls";
        }

        sheetLists.keySet().forEach((mapKey) -> {
            Sheet sheet = workbook.createSheet(mapKey.toString());
            putHeaders(new ArrayList<>(sheetLists.get(mapKey).get(0).keySet()), sheet, workbook);
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

    private void putCellValues(Sheet sheet, List<Map<String, String>> cellValues) {
        int rowNum = 1;
        int cellNum;
        Row row;
        for (Map<String, String> objectValues : cellValues) {
            cellNum = 0;
            row = sheet.createRow(rowNum++);
            for (String parameter : objectValues.values()) {

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
