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
public class Address implements ValueObject, Serializable {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long addressId;

    private String street;
    private String city;
    private String country;
    @Column(columnDefinition = "varbinary(1024)")
    private House house;

    @Override
    public List<String> getValues() {
        List<String> values = new ArrayList<>();
        values.add(this.addressId.toString());
        values.add(this.street);
        values.add(this.city);
        values.add(this.country);
        values.addAll(this.house.getValues());

        return values;
    }

    @Override
    public List<String> getHeaders() {
        List<String> headers = new ArrayList<>();
        headers.add("Address ID");
        headers.add("Street");
        headers.add("City");
        headers.add("Country");
        headers.addAll(this.house.getHeaders());

        return headers;
    }
}
