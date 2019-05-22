package com.example.demo.entity;

import javax.persistence.*;
import java.util.Set;

@Entity
public class Facture {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long Id;

    @ManyToOne
    private Client client;

    @OneToMany(mappedBy = "facture")
    private Set<LigneFacture> ligneFactures;


    public Long getId() {
        return Id;
    }

    public void setId(Long id) {
        Id = id;
    }

    public Client getClient() {
        return client;
    }

    public void setClient(Client client) {
        this.client = client;
    }

    public Set<LigneFacture> getLigneFactures() {
        return ligneFactures;
    }

    public void setLigneFactures(Set<LigneFacture> ligneFactures) {
        this.ligneFactures = ligneFactures;
    }

    public Double getTotal(){
        Double total = 0.0;
        for(LigneFacture ligneFacture : ligneFactures) {
            total = total + ligneFacture.getSousTotal();
        }

        return total;
    }

}
