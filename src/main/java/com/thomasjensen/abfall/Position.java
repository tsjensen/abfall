/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2023 Thomas Jensen
 *
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License, version 3, as published by the Free Software Foundation.
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <https://www.gnu.org/licenses/>.
 */
package com.thomasjensen.abfall;

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

    /** same as {@link #month}, but as object */
    private final YearMonth monthObj;

    /** calendar day number, 1-31, even on months that don't have that many days */
    private int day = -1;

    /** number of rows that make up one day in the final sheet */
    private final int rowsPerDay;

    /** row within the current day (0: Heading, 1: first data cell, ...) */
    private int dayRowIdx;

    private DayOfWeek dayOfWeek = null;



    public Position(final int pYear, final int pMonth, final int pDayRowIdx, final int pRowsPerDay)
    {
        year = pYear;
        monthObj = YearMonth.of(pYear, pMonth);
        dayRowIdx = pDayRowIdx;
        rowsPerDay = pRowsPerDay;
    }



    public void setDay(final int pDay)
    {
        day = pDay;
        if (pDay <= monthObj.lengthOfMonth()) {
            dayOfWeek = LocalDate.of(year, getMonthNumeric1(), pDay).getDayOfWeek();
        }
    }



    public Month getMonth()
    {
        return monthObj.getMonth();
    }



    /**
     * Gets the month-of-year int value.
     * The values are numbered following the ISO-8601 standard, from 1 (January) to 12 (December).
     * @return the month number from 1 to 12
     */
    public int getMonthNumeric1()
    {
        return getMonth().getValue();
    }



    public boolean isOddMonth()
    {
        return getMonthNumeric1() % 2 == 1;
    }



    public boolean isJanuary()
    {
        return monthObj.getMonth() == Month.JANUARY;
    }



    public boolean isDecember()
    {
        return monthObj.getMonth() == Month.DECEMBER;
    }



    public int getDayRowIdx()
    {
        return dayRowIdx;
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
