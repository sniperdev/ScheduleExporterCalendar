package com.sniperdev;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class PdfToXlsConverter {

    public PdfToXlsConverter() throws IOException {

        //Tutaj ładujemy nasz plik pdf
        File file = new File("input.pdf");
        PDDocument pdfDocument = Loader.loadPDF(file);

        //Tutaj tworzymy nasz plik excel
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("PDF Data");
        
        //Wyciągamy tekst z pdf
        PDFTextStripper pdfStripper = new PDFTextStripper();
        String text = pdfStripper.getText(pdfDocument);

        System.out.println(text);

        String[] lines = text.split("\\r?\\n");
        int rowNum = 0;
        for (String line : lines) {
            Row row = sheet.createRow(rowNum++);
            String[] columns = line.split("\\s+");
            int colNum = 0;
            for (String column : columns) {
                row.createCell(colNum++).setCellValue(column);
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")){
            workbook.write(fileOut);
        }

        pdfDocument.close();
        workbook.close();
    }
}
