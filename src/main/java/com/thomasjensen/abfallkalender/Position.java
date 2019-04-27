package com.thomasjensen.abfallkalender;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Month;
import java.time.YearMonth;


/**
 * Describes one cell to be formatted.
 */
public class Position
{
    /** calendar year */
    private final int year;

    /** calendar month as number from 1-12 */
    private final int month;

    /** same as {@link #month}, but as object */
    private final YearMonth monthObj;

    /** calendar day number, 1-31, even on months that don't have that many days */
    private int day = -1;

    /** number of rows that make up one day in the final sheet */
    private final int rowsPerDay;

    /** row within the current day (0: Heading, 1: first data cell, ...) */
    private int dayRowIdx = -1;

    private DayOfWeek dayOfWeek = null;



    public Position(final int pYear, final int pMonth, final int pDayRowIdx, final int pRowsPerDay)
    {
        year = pYear;
        month = pMonth;
        monthObj = YearMonth.of(pYear, month);
        dayRowIdx = pDayRowIdx;
        rowsPerDay = pRowsPerDay;
    }



    public void setDay(final int pDay)
    {
        day = pDay;
        if (pDay <= monthObj.lengthOfMonth()) {
            dayOfWeek = LocalDate.of(year, month, pDay).getDayOfWeek();
        }
    }



    public Month getMonth()
    {
        return monthObj.getMonth();
    }



    public boolean isOddMonth()
    {
        return month % 2 == 1;
    }



    public boolean isJanuary()
    {
        return monthObj.getMonth() == Month.JANUARY;
    }



    public boolean isDecember()
    {
        return monthObj.getMonth() == Month.DECEMBER;
    }



    public boolean isFirstRowOfDay()
    {
        return dayRowIdx == 0;
    }



    public boolean isLastRowOfDay()
    {
        return dayRowIdx == rowsPerDay - 1;
    }



    public boolean dayExists()
    {
        return day <= monthObj.lengthOfMonth();
    }



    public boolean isDay1()
    {
        return day == 1;
    }



    public boolean isDay31()
    {
        return day == 31;
    }



    public boolean isSunday()
    {
        return dayOfWeek == DayOfWeek.SUNDAY;
    }



    public DayOfWeek getDayOfWeek()
    {
        return dayOfWeek;
    }



    public int getDay()
    {
        return day;
    }
}
