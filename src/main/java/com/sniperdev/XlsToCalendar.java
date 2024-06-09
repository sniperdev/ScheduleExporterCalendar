package com.sniperdev;

import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import com.google.api.services.calendar.Calendar;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsToCalendar {
    private final CalendarEventManager calendarEventManager;
    private final String month;

    public XlsToCalendar(String filePath, String month) throws IOException, InvalidFormatException, GeneralSecurityException {
        this.month = month;

        CalendarAuthenticator authenticator = new CalendarAuthenticator();
        Calendar service = authenticator.authenticate();
        this.calendarEventManager = new CalendarEventManager(service);

        File file = new File(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Table 2");

        int rows = sheet.getPhysicalNumberOfRows();
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();

        DataFormatter formatter = new DataFormatter();

        String[][] data = new String[rows][columns];

        //Wypełnianie tablicy danymi
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                data[i][j] = formatter.formatCellValue(sheet.getRow(i).getCell(j));
            }
        }

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

        String[][] calendar = new String[2][rows-3];
        for (int r = 4; r < columns; r++) {
            if (!data[7][r].isEmpty()){
                CalendarEventManager.addEvent(getStartDate(r-3, data[7][r]));
            }

        }
    }

    public String getDayOfWeek(int day){
        String date;
        if(day<10){
            date = "0" + day + "." + month + ".2024";
        }
        else {
            date = day + "." + month + ".2024";
        }
        DateTimeFormatter formatterTwo = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        LocalDate localDate = LocalDate.parse(date, formatterTwo);
        return localDate.getDayOfWeek().toString();
    }

    public ShiftTime getStartDate(int date, String shift){
        ShiftTime shiftTime = new ShiftTime();
        String dayOfWeek = getDayOfWeek(date);
        System.out.println(dayOfWeek);
        String newDate;

        if(date<10){
            newDate = "0"+date;
        }
        else {
            newDate = String.valueOf(date);
        }

        System.out.println(newDate);
        if(dayOfWeek.equals("SATURDAY") || dayOfWeek.equals("SUNDAY")){
            switch (shift){
                case "1":
                    shiftTime.setStartDate("2024-" + month + "-" + newDate + "T07:00:00+02:00");
                    shiftTime.setEndDate("2024-" + month + "-" + newDate + "T15:00:00+02:00");
                    break;
                case "2":
                    shiftTime.setStartDate("2024-" + month + "-" + newDate + "T15:00:00+02:00");
                    shiftTime.setEndDate("2024-" + month + "-" + newDate + "T22:00:00+02:00");
                    break;
            }
        }
        else {
            switch (shift){
                case "R":
                    shiftTime.setStartDate("2024-" + month + "-" + newDate + "T06:00:00+02:00");
                    shiftTime.setEndDate("2024-" + month + "-" + newDate + "T16:00:00+02:00");
                    break;
                case "1":
                    shiftTime.setStartDate("2024-" + month + "-" + newDate + "T08:00:00+02:00");
                    shiftTime.setEndDate("2024-" + month + "-" + newDate + "T16:00:00+02:00");
                    break;
                case "2":
                    shiftTime.setStartDate("2024-" + month + "-" + newDate + "T16:00:00+02:00");
                    shiftTime.setEndDate("2024-" + month + "-" + newDate + "T22:00:00+02:00");
                    break;
            }
        }
        System.out.println(shiftTime.getStartDate() + " " + shiftTime.getEndDate());
        return shiftTime;
    }
}
