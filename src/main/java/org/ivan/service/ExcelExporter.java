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
import java.util.Dictionary;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.ivan.model.Person;
import org.springframework.beans.factory.annotation.Autowired;
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

    public List<String> getHeaders(Class classDefinition) {

        List<String> fieldNames = new ArrayList<>();

        for (Field field : classDefinition.getDeclaredFields()) {
            fieldNames.add(field.getName());
        }

        return fieldNames;
    }

    public <T extends Object> List<String> getHeadersFromGetMethods(T obj) {

        List<String> fieldNames = getHeaders(obj.getClass());
        PropertyDescriptor propertyDescriptor;//Use for getters and setters from property
        Method currentGetMethod;//Current get method from current property
        List<String> methodNames = new ArrayList<>();
        for (String fieldName : fieldNames) {
            try {
                propertyDescriptor = new PropertyDescriptor(fieldName, obj.getClass());
                currentGetMethod = propertyDescriptor.getReadMethod();
                //values.add(currentGetMethod.invoke(obj).toString());

                if (currentGetMethod.invoke(obj).toString().contains("=")) {
                    methodNames.addAll(getHeadersFromGetMethods(currentGetMethod.invoke(obj)));
                } else {
                    methodNames.add(currentGetMethod.getName().replace("get", ""));
                }

                //if(currentGetMethod.invoke(obj).getClass().getSuperclass() != null)
            } catch (IntrospectionException | IllegalAccessException | IllegalArgumentException | InvocationTargetException ex) {
                Logger.getLogger(ExcelExporter.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return methodNames;

    }

    public <T extends Object> List<String> getValuesRecursive(T obj) {

        List<String> fieldNames = getHeaders(obj.getClass());
        PropertyDescriptor propertyDescriptor;//Use for getters and setters from property
        Method currentGetMethod;//Current get method from current property
        List<String> values = new ArrayList<>();
        for (String fieldName : fieldNames) {
            try {
                propertyDescriptor = new PropertyDescriptor(fieldName, obj.getClass());//property (fieldName) from class (classDefinition)
                currentGetMethod = propertyDescriptor.getReadMethod();//gets the definition of the get method

                if (currentGetMethod.invoke(obj).toString().contains("=")) {//obj.get<Parameter>
                    values.addAll(getValuesRecursive(currentGetMethod.invoke(obj)));//Gets all values from getters
                } else {
                    values.add(currentGetMethod.invoke(obj).toString());//if its only 1 value
                }

            } catch (IntrospectionException | IllegalAccessException | IllegalArgumentException | InvocationTargetException ex) {
                Logger.getLogger(ExcelExporter.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return values;
    }
// Pattern for generic type usage
//    public <T extends Object> T getObject() {
//        return null;
//    }

    public <T extends Object> File createExcel(List<T> objects, String fileName, String sheetName) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);

        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);

        int cellNum = 0;
        int rowNum = 0;
        List<String> headers = getHeadersFromGetMethods(objects.get(0));
        Row row = sheet.createRow(rowNum++);
        for (int i = 0; i < headers.size(); i++) {
            String cup = headers.get(i);
            Cell cell = row.createCell(cellNum++);
            cell.setCellValue(cup);
            cell.setCellStyle(style);
        }
        for (T obj : objects) {

            cellNum = 0;
            List<String> parameters = getValuesRecursive(obj);
            row = sheet.createRow(rowNum++);
            for (String parameter : parameters) {

                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(parameter);
                sheet.autoSizeColumn(cellNum);
            }

        }

        try {
            FileOutputStream out
                    = new FileOutputStream(new File(fileName));
            workbook.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    public <T extends Object, E extends Object> File createExcelFromMap(Map<T, List<E>> objects, String fileName) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        for (T mapKey : objects.keySet()) {

            HSSFSheet sheet = workbook.createSheet(mapKey.toString());
            int cellNum = 0;
            int rowNum = 0;
            List<String> headers = getHeadersFromGetMethods(objects.get(mapKey).get(0));

            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < headers.size(); i++) {
                String cup = headers.get(i);
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(cup);
                cell.setCellStyle(style);
            }

            for (E object : objects.get(mapKey)) {
                cellNum = 0;
                List<String> parameters = getValuesRecursive(object);
                row = sheet.createRow(rowNum++);
                for (String parameter : parameters) {

                    Cell cell = row.createCell(cellNum++);
                    cell.setCellValue(parameter);
                    sheet.autoSizeColumn(cellNum);
                }
            }
        }

        File file = new File(fileName + ".xls");
        try {
            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
            out.close();

        } catch (FileNotFoundException e) {
        } catch (IOException e) {
        }

        return null;
    }

}
