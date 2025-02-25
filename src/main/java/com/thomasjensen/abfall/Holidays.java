/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2025 Thomas Jensen
 *
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License, version 3, as published by the Free Software Foundation.
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <https://www.gnu.org/licenses/>.
 */
package com.thomasjensen.abfall;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.Enumeration;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * Knows public holidays from the properties file, and can tell whether a day is a holiday.
 */
public final class Holidays
{
    private static final Pattern DATE_PATTERN = Pattern.compile("(\\d{4})-(\\d\\d?)-(\\d\\d?)");

    private final Properties holidays;



    public Holidays()
    {
        holidays = readPropertyFile();
    }



    private Properties readPropertyFile()
    {
        Properties result = new Properties();
        try (InputStream is = getClass().getResourceAsStream("feiertage.properties")) {
            result.load(new InputStreamReader(is, StandardCharsets.UTF_8));
        }
        catch (IOException e) {
            throw new RuntimeException("failed to load feiertage.properties", e);
        }
        return result;
    }



    public boolean isHoliday(final Position pPos)
    {
        boolean result = false;
        for (final Enumeration<?> h = holidays.propertyNames(); h.hasMoreElements(); ) {
            final String holiday = (String) h.nextElement();
            final Matcher m = DATE_PATTERN.matcher(holidays.getProperty(holiday));
            if (m.matches()) {
                int month = Integer.parseInt(m.group(2));
                int day = Integer.parseInt(m.group(3));
                if (month == pPos.getMonth().getValue() && day == pPos.getDay()) {
                    result = true;
                    break;
                }
            }
            else {
                throw new IllegalStateException("Did not understand date in feiertage.properties: " + holiday);
            }
        }
        return result;
    }
}
