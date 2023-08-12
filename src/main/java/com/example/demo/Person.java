package com.example.demo;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class Person {

    private String nom;
    private String prenom;
    private Integer anneeNaissance;

}
