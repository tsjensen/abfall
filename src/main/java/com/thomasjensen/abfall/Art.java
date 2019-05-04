package com.thomasjensen.abfall;

import java.awt.Color;


/**
 * Trash category. Order indicates importance for placement on chart.
 */
public enum Art
{
    /** "Gartenabfallsammlung" */
    Gartenabfall("GA", new Color(0, 176, 80)),

    /** "Papiermüll" */
    Papier("P", new Color(149, 179, 215)),

    /** "Gelber Sack" */
    GelberSack("GS", new Color(255, 255, 102)),

    /** "Restmüll" */
    Rest("Rest", new Color(207, 121, 119)),

    /** "Biomüll" */
    Bio("Bio", new Color(196, 215, 155)),

    /** "Schadstoffmobil" Location 1 */
    Schadstoff1("SM1", new Color(250, 192, 146)),

    /** "Schadstoffmobil" Location 2 */
    Schadstoff2("SM2", new Color(250, 192, 146)),

    /** "Schadstoffmobil" Location 3 */
    Schadstoff3("SM3", new Color(250, 192, 146)),

    /** "Schadstoffmobil" Location 4 */
    Schadstoff4("SM4", new Color(250, 192, 146));

    //

    private final String abbrev;

    private final Color color;



    private Art(final String pAbbrev, final Color pColor)
    {
        abbrev = pAbbrev;
        color = pColor;
    }



    public String getKuerzel()
    {
        return abbrev;
    }



    public Color getColor()
    {
        return color;
    }
}
