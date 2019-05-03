package com.thomasjensen.abfall;

/**
 * TODO Purpose of this type, in one sentence, ending with a dot. <p>Further, arbitrarily elaborate description. HTML
 * allowed.
 *
 * @author Thomas Jensen
 */
public class Termin
    implements Comparable<Termin>
{
    private final int monat;

    private final int tag;

    private final Art art;



    public Termin(final int pMonat, final int pTag, final Art pArt)
    {
        monat = pMonat;
        tag = pTag;
        if (pArt == null) {
            throw new IllegalArgumentException("pArt must not be null");
        }
        art = pArt;
    }



    public int getMonat()
    {
        return monat;
    }



    public int getTag()
    {
        return tag;
    }



    public Art getArt()
    {
        return art;
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

        Termin termin = (Termin) pOther;

        if (monat != termin.monat) {
            return false;
        }
        if (tag != termin.tag) {
            return false;
        }
        if (art != termin.art) {
            return false;
        }

        return true;
    }



    @Override
    public int hashCode()
    {
        int result = monat;
        result = 31 * result + tag;
        result = 31 * result + art.hashCode();
        return result;
    }



    @Override
    public int compareTo(final Termin pOther)
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
        else if (art.ordinal() > pOther.art.ordinal()) {
            result = 1;
        }
        else if (art.ordinal() < pOther.art.ordinal()) {
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
        sb.append(": ");
        sb.append(art.getKuerzel());
        sb.append('}');
        return sb.toString();
    }
}
