package com.thomasjensen.abfallkalender;

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
