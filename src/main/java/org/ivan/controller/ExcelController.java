/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.controller;

import java.io.File;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import org.ivan.model.Person;
import org.ivan.service.ExcelExporter;
import org.ivan.service.PersonService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

/**
 *
 * @author ivan
 */
@RestController
public class ExcelController {

    @Autowired
    PersonService personService;

    @Autowired
    ExcelExporter excelExporter;

    @RequestMapping(method = RequestMethod.GET)
    public Collection<Person> getAllPeople() {
        return personService.getAllUsers();
    }

    @RequestMapping(method = RequestMethod.POST, consumes = MediaType.APPLICATION_JSON_VALUE)
    public Person addUser(@RequestBody Person person) {
        return personService.addPerson(person);
    }

    @RequestMapping(value = "/export/{parameter}", method = RequestMethod.POST, consumes = MediaType.APPLICATION_JSON_VALUE)
    public File export(@RequestBody String filePath, @PathVariable("parameter") String parameter) {
        File filepath = new File(filePath);
        File path = new File(filepath.getAbsolutePath().replace(filepath.getName(), ""));
        if (!path.exists()) {
            path.mkdir();
        }
        Map<String, List<Map<String, String>>> sheets = excelExporter.getSheetListsByFieldName(new ArrayList<>(personService.getAllUsers()), parameter);
        return excelExporter.createExcelFromSheetMap(sheets, filepath.getAbsolutePath());
    }
}
