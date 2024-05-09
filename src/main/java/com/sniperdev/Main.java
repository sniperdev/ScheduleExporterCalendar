package com.sniperdev;


import com.google.api.services.calendar.Calendar;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.security.GeneralSecurityException;

public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException, GeneralSecurityException {
//        try{
//            CalendarAuthenticator authenticator = new CalendarAuthenticator();
//
//            Calendar service = authenticator.authenticate();
//
//            CalendarEventManager manager = new CalendarEventManager(service);
//
//            manager.listEvents();
//            manager.addEvent();
//        } catch (IOException | GeneralSecurityException e) {
//            e.printStackTrace();
//        }
        XlsToCalendar xlsToCalendar = new XlsToCalendar();
    }
}