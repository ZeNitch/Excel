package org.ivan;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.ivan.model.Address;
import org.ivan.model.Gender;
import org.ivan.model.House;
import org.ivan.model.Person;
import org.ivan.repository.AddressRepository;
import org.ivan.repository.HouseRepository;
import org.ivan.repository.PersonRepository;
import org.ivan.service.ExcelExporter;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelApplicationTests {

    @Autowired
    private PersonRepository personRepository;

    @Autowired
    private AddressRepository addressRepository;

    @Autowired
    private ExcelExporter excelExporter;

    @Autowired
    private HouseRepository houseRepository;

    @Test
    public void contextLoads() throws IOException, InvalidFormatException {
        House h1 = new House();
        h1.setType("Kalkan");
        houseRepository.save(h1);
        House h2 = new House();
        h2.setType("Test");
        houseRepository.save(h2);

        Address a1 = new Address();
        a1.setCity("Ruse");
        a1.setCountry("BG");
        a1.setStreet("Borisova");
        a1.setHouse(h1);
        addressRepository.save(a1);

        Address a2 = new Address();
        a2.setCity("Ruse");
        a2.setCountry("BG");
        a2.setStreet("Borisova");
        a2.setHouse(h2);
        addressRepository.save(a2);

        Address a3 = new Address();
        a3.setCity("Ruse");
        a3.setCountry("BG");
        a3.setStreet("Borisova");
        a3.setHouse(h2);
        addressRepository.save(a3);

        Address a4 = new Address();
        a4.setCity("Ruse");
        a4.setCountry("BG");
        a4.setHouse(h1);
        a4.setStreet("Borisovaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");
        addressRepository.save(a4);

        Person cupPerson = new Person(a1);
        cupPerson.setPersonName("Pesho");
        cupPerson.setGender(Gender.MALE);
        //cupPerson.setAddress(a1);
        personRepository.save(cupPerson);

        Person c1 = new Person(a2);
        c1.setPersonName("Pesha");
        c1.setGender(Gender.FEMALE);
        //c1.setAddress(a2);
        personRepository.save(c1);

        Person c2 = new Person(a3);
        c2.setPersonName("Pesho");
        c2.setGender(Gender.MALE);
        //c2.setAddress(a3);
        personRepository.save(c2);

        Person c3 = new Person(a4);
        c3.setPersonName("Pesho");
        c3.setGender(Gender.FEMALE);
        //c3.setAddress(a4);
        personRepository.save(c3);

        List<Person> personArray = personRepository.findAll();
        File file = excelExporter.createExcel(personRepository.findAll(), "exceltest.xlsx", Gender.MALE.toString());

        Map<String, List<Person>> sheetMap = new HashMap<>();

        for (Person person : personArray) {
            if (!sheetMap.containsKey(person.getAddress().getHouse().getType())) {
                sheetMap.put(person.getAddress().getHouse().getType(), new ArrayList<>());
            }
            sheetMap.get(person.getAddress().getHouse().getType()).add(person);
        }
        Map<Long, List<Person>> sheetMapTest = new HashMap<>();

        for (Person person : personArray) {
            if (!sheetMapTest.containsKey(person.getAddress().getHouse().getHouseId())) {
                sheetMapTest.put(person.getAddress().getHouse().getHouseId(), new ArrayList<>());
            }
            sheetMapTest.get(person.getAddress().getHouse().getHouseId()).add(person);
        }
        File fileTestTwo = excelExporter.createExcelFromMap(sheetMap, "myFile2.xlsx");
        fileTestTwo = excelExporter.createExcelFromMap(sheetMapTest, "myFile2");
        fileTestTwo = excelExporter.createExcelFromMap(sheetMap, "myFile2.xls");
    }

}
