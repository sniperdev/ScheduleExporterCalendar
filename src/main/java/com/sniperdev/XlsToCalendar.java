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

        //Tutaj wyciągamy ilość wiersze
        //LEWO PRAWO
        int rows = sheet.getPhysicalNumberOfRows();

        //Tutaj wyciągamy ilość kolumny
        //GÓRA DÓŁ
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
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < columns; c++) {
                if (data[r][c].equals("Igor Spychała")) {
                    rowIndex = r;
                    columnIndex = c;
                    break;
                }
            }
        }

        //Chciałbym połączyć te dwie tablice w jedną 2D tablicę
        String[][] calendar = new String[2][rows-3];
        for (int r = 4; r < columns; r++) {
            System.out.println(r-3 + " " + data[7][r]);
        }
    }
}
