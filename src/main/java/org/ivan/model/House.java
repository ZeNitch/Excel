/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.ivan.model;

import java.io.Serializable;
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
public class House implements Serializable {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long houseId;

    private String type;
}
