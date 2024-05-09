package com.sniperdev;

public class ShiftTime {
    String startDate;
    String endDate;

    public ShiftTime(String startDate, String endDate) {
        this.startDate = startDate;
        this.endDate = endDate;
    }

    public ShiftTime() {

    }

    public void setEndDate(String endDate) {
        this.endDate = endDate;
    }

    public void setStartDate(String startDate) {
        this.startDate = startDate;
    }

    public String getEndDate() {
        return endDate;
    }

    public String getStartDate() {
        return startDate;
    }
}
