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

public class XlsToCalendarSimplyfied {
    private final CalendarEventManager calendarEventManager;

    public XlsToCalendarSimplyfied() throws IOException, InvalidFormatException, GeneralSecurityException {
        CalendarAuthenticator authenticator = new CalendarAuthenticator();
        Calendar service = authenticator.authenticate();
        this.calendarEventManager = new CalendarEventManager(service);

        File file = new File("input.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Table 2");

        DataFormatter formatter = new DataFormatter();
        String[][] data = extractDataFromSheet(sheet, formatter);

        int[] igorLocation = findIgorLocation(data);
        String[][] calendar = new String[2][data.length-3];
        addEventsToCalendar(data, igorLocation[1], calendar);
    }

    private String[][] extractDataFromSheet(XSSFSheet sheet, DataFormatter formatter) {
        int rows = sheet.getPhysicalNumberOfRows();
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();
        String[][] data = new String[rows][columns];

        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                data[i][j] = formatter.formatCellValue(sheet.getRow(i).getCell(j));
            }
        }
        return data;
    }

    private int[] findIgorLocation(String[][] data) {
        for (int r = 0; r < data.length; r++) {
            for (int c = 0; c < data[0].length; c++) {
                if (data[r][c].equals("Igor Spychała")) {
                    return new int[]{r, c};
                }
            }
        }
        return new int[]{-1, -1};
    }

    private void addEventsToCalendar(String[][] data, int startColumn, String[][] calendar) throws IOException {
        for (int r = 4; r < data[0].length; r++) {
            if (!data[7][r].isEmpty()){
                CalendarEventManager.addEvent(getStartDate(r-3, data[7][r]));
            }
        }
    }

    public String getDayOfWeek(int day){
        String date = String.format("%02d.05.2024", day);
        DateTimeFormatter formatterTwo = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        LocalDate localDate = LocalDate.parse(date, formatterTwo);
        return localDate.getDayOfWeek().toString();
    }

    public ShiftTime getStartDate(int date, String shift){
        ShiftTime shiftTime = new ShiftTime();
        String dayOfWeek = getDayOfWeek(date);
        String newDate = String.format("2024-05-%02dT", date);

        if(dayOfWeek.equals("SATURDAY") || dayOfWeek.equals("SUNDAY")){
            shiftTime.setStartDate(newDate + (shift.equals("1") ? "07:00:00+02:00" : "15:00:00+02:00"));
            shiftTime.setEndDate(newDate + (shift.equals("1") ? "15:00:00+02:00" : "22:00:00+02:00"));
        }
        else {
            shiftTime.setStartDate(newDate + (shift.equals("R") ? "06:00:00+02:00" : (shift.equals("1") ? "08:00:00+02:00" : "16:00:00+02:00")));
            shiftTime.setEndDate(newDate + (shift.equals("R") ? "16:00:00+02:00" : (shift.equals("1") ? "16:00:00+02:00" : "22:00:00+02:00")));
        }
        return shiftTime;
    }
}