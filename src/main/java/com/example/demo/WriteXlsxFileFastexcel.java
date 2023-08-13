package com.example.demo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class WriteXlsxFileFastexcel {

    public static void main(String[] args) throws IOException {
        write("src/main/resources/writeFastexcel.xlsx");
    }

    private static void write(String cheminFichier) throws IOException {
        log.info("DÉBUT");
        LocalDateTime debut = LocalDateTime.now();
        try (FileOutputStream outputStream = new FileOutputStream(cheminFichier);
                Workbook wb = new Workbook(outputStream, "toto", null);) {
            Worksheet worksheet = wb.newWorksheet("Sheet 1");
            worksheet.width(0, 25);
            worksheet.width(1, 15);
            worksheet.range(0, 0, 0, 1).style().fontName("Arial").fontSize(16).bold().fillColor("3366FF").set();
            worksheet.value(0, 0, "Name");
            worksheet.value(0, 1, "Age");
            for (int i = 1; i < 500_000; i++) {
                worksheet.value(i, 0, "John Smith");
                worksheet.value(i, 1, 20);
            }
            LocalDateTime fin = LocalDateTime.now();
            log.info("FIN : " + Duration.between(debut, fin).toMillis() + "ms");
        }
    }

}
