/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.service;

import java.util.Collection;
import org.ivan.model.Person;
import org.ivan.repository.PersonRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

/**
 *
 * @author ivan
 */
@Service
public class PersonService {

    @Autowired
    PersonRepository personRepository;

    public Collection<Person> getAllUsers() {
        return personRepository.findAll();
    }

    public Person addPerson(Person person) {
        return personRepository.save(person);
    }
}
