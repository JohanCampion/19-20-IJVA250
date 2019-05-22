package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Set;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();

        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        writer.println("Id" + ";" + "Nom" + ";" + "Prenom" + ";" + "Date de Naissance" + ";" + "Age");

        for( Client client : allClients ){
            writer.println( client.getId() + ";" + client.getNom() + ";\"" + client.getPrenom() + "\";" +
                    client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")) + ";" + (now.getYear() - client.getDateNaissance().getYear()));
        }

    }

    @GetMapping("/clients/xlsx")
    public void clientXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        Cell cellNom = headerRow.createCell(1);
        Cell cellPrenom = headerRow.createCell(2);
        Cell cellDn = headerRow.createCell(3);
        Cell cellAge = headerRow.createCell(4);
        cellId.setCellValue("Id");
        cellNom.setCellValue("Nom");
        cellPrenom.setCellValue("Prenom");
        cellDn.setCellValue("Date de naissance");
        cellAge.setCellValue("Age");


        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        int i = 1;
        for( Client client : allClients ){
            Row newRow = sheet.createRow(i);
            newRow.createCell(0).setCellValue(client.getId());
            newRow.createCell(1).setCellValue(client.getNom());
            newRow.createCell(2).setCellValue(client.getPrenom());
            newRow.createCell(3).setCellValue(client.getDateNaissance().toString());
            newRow.createCell(4).setCellValue((now.getYear() - client.getDateNaissance().getYear()));
            i++;
        }


        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{idClient}/factures/xlsx")
    public void getFacture(@PathVariable String idCLient, HttpServletRequest request, HttpServletResponse response) {
//
//        response.setContentType("text/xlsx");
//        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
//
//        Workbook workbook = new XSSFWorkbook();
//
//        Row headerRow = sheet.createRow(0);
//
//        Cell cellId = headerRow.createCell(0);
//        Cell cellNom = headerRow.createCell(1);
//        Cell cellPrenom = headerRow.createCell(2);
//        Cell cellDn = headerRow.createCell(3);
//        Cell cellAge = headerRow.createCell(4);
//        cellId.setCellValue("Id");
//        cellNom.setCellValue("Nom");
//        cellPrenom.setCellValue("Prenom");
//        cellDn.setCellValue("Date de naissance");
//        cellAge.setCellValue("Age");
//
//
//        List<Facture> factureClient = factureService.findFactureByClient(idCLient);
//
//        for(Facture facture : factureClient){
//            Set<LigneFacture> ligneFactures = facture.getLigneFactures();
//
//            for(LigneFacture ligneFact : ligneFactures) {
//
//                Sheet sheet = workbook.createSheet("facture n°" + facture.getId());
//
//                Row factureHeaderRow = sheet.createRow(0);
//
//                Cell cellDesignation = factureHeaderRow.createCell(0);
//                cellDesignation.setCellValue("Désignation");
//
//                Cell Quantite = factureHeaderRow.createCell(1);
//                cellDesignation.setCellValue("Quantité");
//
//                Cell cellPrix = factureHeaderRow.createCell(2);
//                cellDesignation.setCellValue("Prix");
//
//
//            }
//
//        }


    }

    @GetMapping("/factures/xlsx")
    public void getAllFactures(HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        Workbook workbook = new XSSFWorkbook();

        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        for(Client client : allClients){

            Sheet sheetClient = workbook.createSheet(client.getNom());
            sheetClient.createRow(0).createCell(0).setCellValue(client.getNom());
            sheetClient.createRow(1).createCell(0).setCellValue(client.getPrenom());
            sheetClient.createRow(2).createCell(0).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));


            List<Facture> factures = factureService.findFactureByClient(client.getId());

            //création des pages de factures
            for(Facture facture : factures){

                Sheet sheetFacture = workbook.createSheet("Facture" + facture.getId());
                Row headerRow = sheetFacture.createRow(0);
                headerRow.createCell(0).setCellValue("nom de l'article");
                headerRow.createCell(1).setCellValue("quantité");
                headerRow.createCell(2).setCellValue("prix unitaire");
                headerRow.createCell(3).setCellValue("sous-total");

                //création des lignes de chaque factures
                int i = 1;
                for(LigneFacture ligneFacture : facture.getLigneFactures()){
                    Row newRow = sheetFacture.createRow(i);
                    newRow.createCell(0).setCellValue(ligneFacture.getArticle().getLibelle());
                    newRow.createCell(1).setCellValue(ligneFacture.getQuantite());
                    newRow.createCell(2).setCellValue(ligneFacture.getArticle().getPrix());
                    newRow.createCell(3).setCellValue(ligneFacture.getSousTotal());
                    i++;
                }

                //Création ligne du Total de la facture
                Row totalRow = sheetFacture.createRow(i);
                Cell totalCell = totalRow.createCell(2);
                totalCell.setCellValue("TOTAL");
                Cell totalCell2 = totalRow.createCell(3);
                totalCell2.setCellValue(facture.getTotal());

                //Ajout mise en page

                CellStyle cellStyle  = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                font.setColor(IndexedColors.RED.getIndex());

                cellStyle.setBorderBottom(BorderStyle.MEDIUM);
                cellStyle.setBorderLeft(BorderStyle.MEDIUM);
                cellStyle.setBorderTop(BorderStyle.MEDIUM);
                cellStyle.setBorderRight(BorderStyle.MEDIUM);

                cellStyle.setFont(font);

                totalCell.setCellStyle(cellStyle);
                totalCell2.setCellStyle(cellStyle);
            }

        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }



}
