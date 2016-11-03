/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import lombok.Data;

/**
 *
 * @author ivan
 */
@Entity
@Data
public class Person implements ValueObject, Serializable {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long personId;

    private String personName;
    private Gender gender;
    @Column(columnDefinition = "varbinary(1024)")
    private Address address;

    public Person(Address address) {
        this.address = address;
    }

    public Person() {
    }

    @Override
    public List<String> getValues() {
        List<String> values = new ArrayList<String>();
        values.add(this.personId.toString());
        values.add(this.personName);
        values.add(this.gender.toString());
        values.addAll(this.address.getValues());

        return values;
    }

    @Override
    public List<String> getHeaders() {

        List<String> getHeaders = new ArrayList<String>();
        getHeaders.add("Person ID");
        getHeaders.add("Person Name");
        getHeaders.add("Gender");
        getHeaders.addAll(this.address.getHeaders());

        return getHeaders;
    }

}
