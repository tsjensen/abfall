package com.thomasjensen.abfall;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.DayOfWeek;
import java.time.format.TextStyle;
import java.util.Calendar;
import java.util.List;
import java.util.Locale;
import java.util.SortedMap;
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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main
{
    private static final int YEAR = 2019;

    private static final Locale LOCALE = Locale.GERMAN;

    private static final File ICS_FILE = new File("_support/Abfallkalender " + YEAR + ".ics");



    private static String getDayOfWeekDisplay(final DayOfWeek pDayOfWeek)
    {
        return pDayOfWeek.getDisplayName(TextStyle.SHORT_STANDALONE, LOCALE);
    }



    public static void main(String[] args)
        throws IOException
    {
        final ICalendar cal = Biweekly.parse(ICS_FILE).first();
        final List<VEvent> events = cal.getEvents();

        System.out.println("read " + events.size() + " termine");

        final SortedMap<Zeitpunkt, SortedSet<Art>> termine = new TreeMap<>();
        for (final VEvent event : events) {
            final ICalDate datum = event.getProperty(DateStart.class).getValue();
            final int monat = datum.getMonth() + 1;
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

            final Zeitpunkt z = new Zeitpunkt(monat, tag);
            SortedSet<Art> arten = termine.computeIfAbsent(z, k -> new TreeSet<>());
            arten.add(art);
            if (arten.size() > 2) {
                throw new IllegalStateException("too many events per day on " + z);
            }
        }

        // for (final Termin termin : termine) {
        //     System.out.println(termin);
        // }

        final int linesPerMonth = 3;
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final CellStyleFactory cellStyleFactory = new CellStyleFactory(workbook);
        OutputStream fileOut = new FileOutputStream("Abfallkalender " + YEAR + " test.xlsx");
        Sheet sheet1 = workbook.createSheet("Abholtermine");

        int xlRowNum = 0;
        Row xlRow = sheet1.createRow(xlRowNum++);
        xlRow.setHeightInPoints(35.25f);
        Cell headingCell = xlRow.createCell(1);
        headingCell.setCellValue("Abholtermine");
        headingCell.setCellStyle(cellStyleFactory.heading());
        xlRow = sheet1.createRow(xlRowNum++);
        xlRow.setHeightInPoints(29.25f);
        Cell excelCell = xlRow.createCell(0);
        excelCell.setCellValue(YEAR);
        excelCell.setCellStyle(cellStyleFactory.yearHeading());
        sheet1.setColumnWidth(0, 4064);  // column width 13.86
        for (int i = 1; i <= 31; i++) {
            Cell c = xlRow.createCell(i);
            c.setCellValue(i);
            c.setCellStyle(cellStyleFactory.columnHeading(i));
            sheet1.setColumnWidth(i, 1340);  // column width 4.57
        }

        for (int xlRowOffset = 0; xlRowOffset < linesPerMonth * 12; xlRowOffset++) {
            xlRow = sheet1.createRow(xlRowNum++);
            final int month = (xlRowOffset / linesPerMonth) + 1;
            final Position pos = new Position(YEAR, month, xlRowOffset % linesPerMonth, linesPerMonth);

            excelCell = xlRow.createCell(0);
            if (pos.isFirstRowOfDay()) {
                excelCell.setCellValue(" " + pos.getMonth().getDisplayName(TextStyle.FULL, LOCALE));
            }
            excelCell.setCellStyle(cellStyleFactory.monthHeading(pos));

            for (int day = 1; day <= 31; day++) {
                pos.setDay(day);
                final SortedSet<Art> categories = termine.get(new Zeitpunkt(month, day));
                final boolean isDayEmpty = categories == null || categories.isEmpty();
                excelCell = xlRow.createCell(day);

                if (!pos.dayExists()) {
                    excelCell.setBlank();
                    excelCell.setCellStyle(
                        pos.isOddMonth() ? cellStyleFactory.oddRowCentered(pos) : cellStyleFactory.centered(pos));
                }
                else if (isDayEmpty) {
                    if (pos.isFirstRowOfDay()) {
                        excelCell.setCellValue(getDayOfWeekDisplay(pos.getDayOfWeek()));
                    }
                    if (pos.isSunday()) {
                        excelCell.setCellStyle(cellStyleFactory.sunday(pos));
                    }
                    else {
                        excelCell.setCellStyle(
                            pos.isOddMonth() ? cellStyleFactory.oddRowCentered(pos) : cellStyleFactory.centered(pos));
                    }
                }
                else if (pos.isFirstRowOfDay()) {
                    excelCell.setCellValue(getDayOfWeekDisplay(pos.getDayOfWeek()));
                    excelCell.setCellStyle(cellStyleFactory.dayHeading(pos, categories.first()));
                }
                else {
                    if (categories.size() >= xlRowOffset % linesPerMonth) {
                        final Art category = xlRowOffset % linesPerMonth == 1 ? categories.first() : categories.last();
                        excelCell.setCellValue(category.getKuerzel());
                        excelCell.setCellStyle(cellStyleFactory.dayCell(pos, category));
                    }
                    else {
                        excelCell.setBlank();
                        excelCell.setCellStyle(cellStyleFactory.centered(pos));
                    }
                    if (xlRowOffset % linesPerMonth == 2 && categories.size() == 1) {
                        sheet1.addMergedRegion(new CellRangeAddress(xlRowNum - 2, xlRowNum - 1, day, day));
                    }
                }
            }
        }

        mergeMonthNames(sheet1, linesPerMonth);

        sheet1.createRow(xlRowNum++);
        xlRow = sheet1.createRow(xlRowNum++);
        excelCell = xlRow.createCell(1);
        excelCell.setCellValue(
            "Öffnungszeiten Hafen:  Mo. - Fr.:  7 - 12 und 13 - 17 Uhr,   Samstag:  8 - 14 Uhr,   Mo./Di. kein "
                + "Sondermüll");
        excelCell.setCellStyle(cellStyleFactory.noteBig());

        xlRow = sheet1.createRow(xlRowNum++);
        excelCell = xlRow.createCell(1);
        excelCell.setCellValue("Gartenabfallsammlungen:");
        excelCell.setCellStyle(cellStyleFactory.noteBig());
        excelCell = xlRow.createCell(7);
        excelCell.setCellValue("verschiedene Standorte");
        excelCell.setCellStyle(cellStyleFactory.noteSmall());
        excelCell = xlRow.createCell(13);
        excelCell.setCellValue("Schadstoffmobil:");
        excelCell.setCellStyle(cellStyleFactory.noteBig());
        excelCell = xlRow.createCell(17);
        excelCell.setCellValue("ab 1.1.2019 abgeschafft");
        excelCell.setCellStyle(cellStyleFactory.noteSmall());

        xlRow = sheet1.createRow(xlRowNum++);    // for inclusion in print area, and extra notes
        workbook.setPrintArea(
            0,  //sheet index
            0,  //start column
            31, //end column
            0,  //start row
            41  //end row
        );
        sheet1.getPrintSetup().setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);

        workbook.write(fileOut);
        fileOut.close();

        // http://stackoverflow.com/questions/33901/best-icalendar-library-for-java
        // http://sourceforge.net/p/biweekly/wiki/Quick%20Start/
    }



    @SuppressWarnings("SameParameterValue")
    private static void mergeMonthNames(final Sheet pSheet, final int pLinesPerMonth)
    {
        for (int month = Calendar.JANUARY; month <= Calendar.DECEMBER; month++) {
            int xlRowNum = month * pLinesPerMonth + 2;
            pSheet.addMergedRegion(new CellRangeAddress(xlRowNum, xlRowNum + pLinesPerMonth - 1, 0, 0));
        }
    }
}
