package com.thomasjensen.abfall;

/**
 * Trash category. Order indicates importance for placement on chart.
 *
 * @author Thomas Jensen
 */
public enum Art
{
    /** Gartenabfallsammlung */
    Gartenabfall("GA"),

    /** Papiermüll */
    Papier("P"),

    /** Gelber Sack */
    GelberSack("GS"),

    /** Restmüll */
    Rest("Rest"),

    /** Biomüll */
    Bio("Bio"),

    /** Schadstoffmobil Standort 1 */
    Schadstoff1("SM1"),

    /** Schadstoffmobil Standort 2 */
    Schadstoff2("SM2"),

    /** Schadstoffmobil Standort 3 */
    Schadstoff3("SM3"),

    /** Schadstoffmobil Standort 4 */
    Schadstoff4("SM4");

    //

    private final String abbrev;



    private Art(final String pAbbrev)
    {
        abbrev = pAbbrev;
    }



    public String getKuerzel()
    {
        return abbrev;
    }
}
