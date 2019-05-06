/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2019 Thomas Jensen
 *
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License, version 3, as published by the Free Software Foundation.
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <https://www.gnu.org/licenses/>.
 */
package com.thomasjensen.abfall;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.SortedSet;
import java.util.TreeMap;
import java.util.TreeSet;

import biweekly.Biweekly;
import biweekly.ICalendar;
import biweekly.component.VEvent;
import biweekly.property.DateStart;
import biweekly.property.Location;
import biweekly.property.Summary;
import biweekly.util.ICalDate;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// TODO license
// TODO dependency updates
// TODO readme



public class Main
{
    private static final Logger LOG = LogManager.getLogger(Main.class);



    public static void main(final String[] pArgs)
        throws IOException
    {
        new Main().entrypoint(pArgs);
    }



    public void entrypoint(final String[] pArgs)
        throws IOException
    {
        printLicenseHeader();
        final Config config = new CmdLine().parse(pArgs);
        if (config == null) {
            System.exit(0);
        }

        final List<VEvent> events = readIcsFile(config);
        final Map<DayOnCalendar, SortedSet<Art>> termine = groupByDay(events);
        final XSSFWorkbook workbook = new ExcelCreator(config).create(termine);

        writeWorkbook(workbook, config);
    }



    private void printLicenseHeader()
    {
        LOG.info("abfall - convert ICS format trash calendar into a single Excel sheet");
        LOG.info("Copyright (C) 2011-2019 Thomas Jensen");
        LOG.info("This program comes with ABSOLUTELY NO WARRANTY; see LICENSE.md for details.");
        LOG.info("This is free software, and you are welcome to redistribute it under certain conditions; "
            + "see LICENSE.md for details.");
    }



    private Map<DayOnCalendar, SortedSet<Art>> groupByDay(final List<VEvent> pEvents)
    {
        final Map<DayOnCalendar, SortedSet<Art>> result = new TreeMap<>();
        for (final VEvent event : pEvents) {
            final ICalDate datum = event.getProperty(DateStart.class).getValue();
            @SuppressWarnings("deprecation")
            final int monat = datum.getMonth() + 1;
            @SuppressWarnings("deprecation")
            final int tag = datum.getDate();

            final String desc = event.getProperty(Summary.class).getValue().toLowerCase();
            Art art = null;
            if (desc.contains("bio")) {
                art = Art.Bio;
            }
            else if (desc.contains("papier")) {
                art = Art.Papier;
            }
            else if (desc.contains("rest")) {
                art = Art.Rest;
            }
            else if (desc.contains("sack")) {
                art = Art.GelberSack;
            }
            else if (desc.contains("garten")) {
                art = Art.Gartenabfall;
            }
            else if (desc.contains("schadstoff")) {
                // this logic is currently unused because "Schadstoffmobil" was discontinued
                String loc = event.getProperty(Location.class).getValue();
                loc = loc.substring(loc.indexOf("Info: ") + "Info: ".length()).toLowerCase();
                if (loc.contains("some street 1")) {
                    art = Art.Schadstoff1;
                }
                else if (loc.contains("some other street 2")) {
                    art = Art.Schadstoff2;
                }
                else {
                    art = Art.Schadstoff3;
                }
            }

            final DayOnCalendar z = new DayOnCalendar(monat, tag);
            SortedSet<Art> arten = result.computeIfAbsent(z, k -> new TreeSet<>());
            arten.add(art);
            if (arten.size() > 2) {
                throw new IllegalStateException("too many events per day on " + z);
            }
        }

        if (LOG.isDebugEnabled()) {
            for (Map.Entry<DayOnCalendar, SortedSet<Art>> entry : result.entrySet()) {
                for (final Art art : entry.getValue()) {
                    LOG.debug(entry.getKey() + " -> " + art.toString());
                }
            }
        }
        return result;
    }



    private List<VEvent> readIcsFile(final Config pConfig)
        throws IOException
    {
        if (LOG.isInfoEnabled()) {
            LOG.info("Reading ICS file: " + pConfig.getInFileIcs() + " ...");
        }

        // http://sourceforge.net/p/biweekly/wiki/Quick%20Start/
        final ICalendar cal = Biweekly.parse(pConfig.getInFileIcs()).first();
        final List<VEvent> events = cal.getEvents();

        if (LOG.isInfoEnabled()) {
            LOG.info("Parsed " + events.size() + " dates from the ICS file.");
        }
        return events;
    }



    private void writeWorkbook(final XSSFWorkbook pWorkbook, final Config pConfig)
        throws IOException
    {
        try (OutputStream fileOut = new FileOutputStream(pConfig.getOutFileXlsx())) {
            pWorkbook.write(fileOut);
            LOG.info("Generated output file at " + pConfig.getOutFileXlsx());
        }
    }
}
