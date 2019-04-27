package com.thomasjensen.abfallkalender;

/**
 * TODO Purpose of this type, in one sentence, ending with a dot. <p>Further, arbitrarily elaborate description. HTML
 * allowed.
 *
 * @author Thomas Jensen
 */
public class Zeitpunkt
    implements Comparable<Zeitpunkt>
{
    private final int monat;

    private final int tag;



    public Zeitpunkt(final int pMonat, final int pTag)
    {
        monat = pMonat;
        tag = pTag;
    }



    public int getMonat()
    {
        return monat;
    }



    public int getTag()
    {
        return tag;
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

        Zeitpunkt termin = (Zeitpunkt) pOther;
        return compareTo(termin) == 0;
    }



    @Override
    public int hashCode()
    {
        int result = monat;
        result = 31 * result + tag;
        return result;
    }



    @Override
    public int compareTo(final Zeitpunkt pOther)
    {
        int result = 0;
        if (monat > pOther.monat) {
            result = 1;
        }
        else if (monat < pOther.monat) {
            result = -1;
        }
        else if (tag > pOther.tag) {
            result = 1;
        }
        else if (tag < pOther.tag) {
            result = -1;
        }
        return result;
    }



    @Override
    public String toString()
    {
        StringBuilder sb = new StringBuilder();
        sb.append('{');
        if (tag < 10) {
            sb.append('0');
        }
        sb.append(tag);
        sb.append('.');
        if (monat < 10) {
            sb.append('0');
        }
        sb.append(monat);
        sb.append('}');
        return sb.toString();
    }
}
