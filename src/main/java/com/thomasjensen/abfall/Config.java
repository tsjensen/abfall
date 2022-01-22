/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2022 Thomas Jensen
 *
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License, version 3, as published by the Free Software Foundation.
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <https://www.gnu.org/licenses/>.
 */
package com.thomasjensen.abfall;

import java.io.File;
import java.util.Locale;


/**
 * Configuration as passed via the command line.
 */
public class Config
{
    private final int year;

    private final Locale locale;

    private final File inFileIcs;

    private final File outFileXlsx;



    public Config(final int pYear, final Locale pLocale, final File pInFileIcs, final File pOutFileXlsx)
    {
        year = pYear;
        locale = pLocale;
        inFileIcs = pInFileIcs;
        outFileXlsx = pOutFileXlsx;
    }



    public int getYear()
    {
        return year;
    }



    public Locale getLocale()
    {
        return locale;
    }



    public File getInFileIcs()
    {
        return inFileIcs;
    }



    public File getOutFileXlsx()
    {
        return outFileXlsx;
    }
}
