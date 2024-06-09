package com.sniperdev;

import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.println("Witaj w eksporterze pliku Excel do Google Calendar");
        String filePath = "C:\\Users\\igiig\\Documents\\Webstorm Projects\\JAVA\\ScheduleExporterCalendar\\src\\main\\resources\\input.xlsx";
        System.out.println("Podaj numer miesiąca w formacie np. 06: ");
        String month = scanner.nextLine();

        try{
            XlsToCalendar xlsToCalendar = new XlsToCalendar(filePath, month);
        } catch (Exception e){
            System.out.println("Wystąpił błąd podczas przetwarzania pliku: " + e.getMessage());
        }
    }
}