package com.sniperdev;

import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.println("Witaj w eksporterze pliku Excel do Google Calendar");
        System.out.println("Podaj ścieżkę do pliku Excel z twoim grafikiem: ");
        String filePath = scanner.nextLine();

        try{
            XlsToCalendar xlsToCalendar = new XlsToCalendar(filePath);
        } catch (Exception e){
            System.out.println("Wystąpił błąd podczas przetwarzania pliku: " + e.getMessage());
        }
    }
}