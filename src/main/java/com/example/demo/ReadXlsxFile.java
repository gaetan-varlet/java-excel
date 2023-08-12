package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXlsxFile {

    public static void main(String[] args) throws IOException {
        readAllFile("src/main/resources/test.xlsx");
        readOneCell("src/main/resources/test.xlsx");
    }

    private static void readAllFile(String cheminFichier) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(cheminFichier));
                Workbook workbook = new XSSFWorkbook(fis);) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType().equals(CellType.STRING)) {
                        System.out.print(cell.getStringCellValue() + "\t");
                    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                        System.out.print(cell.getNumericCellValue() + "\t");
                    }
                    // les autres types sont ignorés : formules, booléens...
                }
                System.out.println();
            }
        }
    }

    private static void readOneCell(String cheminFichier) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(cheminFichier));
                Workbook workbook = new XSSFWorkbook(fis);) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(1);
            Cell cell = row.getCell(1);
            System.out.println("Contenu de la cellule : " + cell.getRichStringCellValue().getString());
        }
    }

}
