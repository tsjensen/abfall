package com.thomasjensen.abfall;

/**
 * Identifies a day in the resulting calendar.
 */
public class DayOnCalendar
    implements Comparable<DayOnCalendar>
{
    private final int month;

    private final int day;



    public DayOnCalendar(final int pMonat, final int pTag)
    {
        month = pMonat;
        day = pTag;
    }



    public DayOnCalendar(final Position pPos)
    {
        month = pPos.getMonthNumeric1();
        day = pPos.getDay();
    }



    @Override
    public boolean equals(final Object pOther)
    {
        if (this == pOther) {
            return true;
        }
        if (pOther == null || getClass() != pOther.getClass()) {
            return false;
        }

        DayOnCalendar termin = (DayOnCalendar) pOther;
        return compareTo(termin) == 0;
    }



    @Override
    public int hashCode()
    {
        int result = month;
        result = 31 * result + day;
        return result;
    }



    @Override
    public int compareTo(final DayOnCalendar pOther)
    {
        int result = 0;
        if (month > pOther.month) {
            result = 1;
        }
        else if (month < pOther.month) {
            result = -1;
        }
        else if (day > pOther.day) {
            result = 1;
        }
        else if (day < pOther.day) {
            result = -1;
        }
        return result;
    }



    @Override
    public String toString()
    {
        StringBuilder sb = new StringBuilder();
        sb.append('{');
        if (day < 10) {
            sb.append('0');
        }
        sb.append(day);
        sb.append('.');
        if (month < 10) {
            sb.append('0');
        }
        sb.append(month);
        sb.append('}');
        return sb.toString();
    }
}
