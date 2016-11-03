/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
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
public class House implements ValueObject, Serializable {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long houseId;

    private String type;

    @Override
    public List<String> getValues() {
        List<String> values = new ArrayList<>();
        values.add(this.houseId.toString());
        values.add(this.type);

        return values;
    }

    @Override
    public List<String> getHeaders() {
        StringBuilder bobTheBuilder = new StringBuilder();
        List<String> headers = new ArrayList<>();
        headers.add("House ID");
        headers.add("Type");

        return headers;
    }
}
