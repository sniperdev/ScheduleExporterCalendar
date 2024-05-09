package com.sniperdev;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsToCalendar {


    public XlsToCalendar() throws IOException, InvalidFormatException {
        //Zaciągamy plik
        File file = new File("input.xlsx");
        //Tworzymy workbook
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Table 2");

        //Tutaj wyciągamy ilość wierszy
        int rows = sheet.getPhysicalNumberOfRows();

        //Tutaj wyciągamy ilość kolumn
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();

        //Tutaj wyciągamy dane z komórki
        DataFormatter formatter = new DataFormatter();

        //Tworzenie 2D tablicy
        String[][] data = new String[rows][columns];

        //Wypełnianie tablicy danymi
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                data[i][j] = formatter.formatCellValue(sheet.getRow(i).getCell(j));
            }
        }

        //Znalezienie indeksu imienia
        int rowIndex = -1;
        int columnIndex = -1;
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                if (data[i][j].equals("Igor Spychała")) {
                    rowIndex = i;
                    columnIndex = j;
                    break;
                }
            }
        }
        System.out.println("Index: " + columnIndex + " " + rowIndex);
        System.out.println("Data: " + data[rowIndex][columnIndex]);

        int rowIndexDay = -1;
        int columnIndexDay = -1;
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                if (data[i][j].equals("Ilość zmian")) {
                    rowIndexDay = i;
                    columnIndexDay = j;
                    break;
                }
            }
        }
        System.out.println("Index: " + columnIndexDay + " " + rowIndexDay);
        System.out.println("Data: " + data[rowIndexDay][columnIndexDay]);

        //Wypisanie od data to końca wiersza
        for (int i = columnIndexDay+1; i < columns; i++) {
            System.out.println(data[rowIndexDay][i]);
        }

        //Chciałbym połączyć te dwie tablice w jedną 2D tablicę
        String[][] calendar = new String[2][columns-3];
        for (int i = 0; i < columns-3; i++) {
            calendar[0][i] = data[rowIndexDay][i+columnIndexDay+1];
            calendar[1][i] = data[rowIndex][i+columnIndex+1];
        }
        //Wypisanie calendar
        for (int i = 0; i < 2; i++) {
            for (int j = 0; j < columns; j++) {
                System.out.print(calendar[i][j] + " ");
            }
            System.out.println();
        }
    }
}
