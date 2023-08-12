package com.example.demo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteXlsxFile {

    public static void main(String[] args) throws IOException {
        write("src/main/resources/newFile.xlsx");
    }

    private static void write(String cheminFichier) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();) {
            Sheet sheet = workbook.createSheet("Persons");
            sheet.setColumnWidth(0, 6000);
            sheet.setColumnWidth(1, 4000);
            ligneTitres(workbook, sheet);
            ligne1(sheet);
            sheet = workbook.createSheet("Famille");
            writeSheetWithCollection(sheet, getPeople());
            try (FileOutputStream outputStream = new FileOutputStream(cheminFichier);) {
                workbook.write(outputStream);
            }
            System.out.println("Création terminée");
        }
    }

    private static void ligneTitres(Workbook workbook, Sheet sheet) {
        Row header = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);
    }

    private static void ligne1(Sheet sheet) {
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("John Smith");
        cell = row.createCell(1);
        cell.setCellValue(20);
    }

    private static List<Person> getPeople() {
        return List.of(Person.builder().nom("Varlet").prenom("Gaëtan").anneeNaissance(1988).build(),
                Person.builder().nom("Varlet").prenom("Louis").anneeNaissance(2018).build());
    }

    private static void writeSheetWithCollection(Sheet sheet, List<Person> people) {
        int rowCount = 0;
        writeTitle(sheet.createRow(rowCount++));
        for (Person person : people) {
            Row row = sheet.createRow(rowCount++);
            writePerson(person, row);
        }
    }

    private static void writeTitle(Row row) {
        Cell cell = row.createCell(0);
        cell.setCellValue("Nom");
        cell = row.createCell(1);
        cell.setCellValue("Prénom");
        cell = row.createCell(2);
        cell.setCellValue("Année de naissance");
    }

    private static void writePerson(Person person, Row row) {
        Cell cell = row.createCell(0);
        cell.setCellValue(person.getNom());
        cell = row.createCell(1);
        cell.setCellValue(person.getPrenom());
        cell = row.createCell(2);
        cell.setCellValue(person.getAnneeNaissance());
    }

}
