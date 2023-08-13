package com.example.demo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class WriteXlsxFileApachePoi {

    public static void main(String[] args) throws IOException {
        write("src/main/resources/newFileApachePoi.xlsx");
    }

    private static void write(String cheminFichier) throws IOException {
        log.info("DÉBUT");
        LocalDateTime debut = LocalDateTime.now();
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Persons");
            for (int i = 1; i < 500_000; i++) {
                ligne1(sheet, i);
            }
            try (FileOutputStream outputStream = new FileOutputStream(cheminFichier);) {
                workbook.write(outputStream);
            }
            LocalDateTime fin = LocalDateTime.now();
            log.info("FIN : " + Duration.between(debut, fin).toMillis() + "ms");
        }
    }

    private static void ligne1(Sheet sheet, int i) {
        Row row = sheet.createRow(i);
        Cell cell = row.createCell(0);
        cell.setCellValue("John Smith");
        cell = row.createCell(1);
        cell.setCellValue(20);
    }

}
