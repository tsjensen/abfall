/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2020 Thomas Jensen
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
import java.time.DayOfWeek;
import java.time.format.TextStyle;
import java.util.Calendar;
import java.util.Locale;
import java.util.Map;
import java.util.SortedSet;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Create an Excel workbook (our output).
 */
public class ExcelCreator
{
    private static final int ROWS_PER_DAY = 3;

    private final Config config;

    private final XSSFWorkbook workbook;

    private final XSSFSheet sheet;

    private final CellStyleFactory cellStyleFactory;

    private int xlRowNum = 0;



    public ExcelCreator(final Config pConfig)
    {
        config = pConfig;
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Abholtermine");
        cellStyleFactory = new CellStyleFactory(workbook);
    }



    public XSSFWorkbook create(final Map<DayOnCalendar, SortedSet<Art>> pTermine)
    {
        createHeadings();

        for (int month = Calendar.JANUARY; month <= Calendar.DECEMBER; month++) {
            for (int dayRowIdx = 0; dayRowIdx < ROWS_PER_DAY; dayRowIdx++) {
                final Position pos = new Position(config.getYear(), month + 1, dayRowIdx, ROWS_PER_DAY);

                final Row xlRow = sheet.createRow(xlRowNum++);
                addMonthHeadingCell(xlRow, pos);

                for (int day = 1; day <= 31; day++) {
                    pos.setDay(day);

                    SortedSet<Art> categories = pTermine.get(new DayOnCalendar(pos));
                    addDayCell(xlRow, pos, categories);
                }
            }
        }

        mergeMonthNames();
        addNotices();
        addQrCode();
        setPrintSetup();
        setWorkbookProperties();

        return workbook;
    }



    private void createHeadings()
    {
        Row xlRow = sheet.createRow(xlRowNum++);
        xlRow.setHeightInPoints(35.25f);
        Cell headingCell = xlRow.createCell(1);
        headingCell.setCellValue("Abholtermine");
        headingCell.setCellStyle(cellStyleFactory.heading());

        xlRow = sheet.createRow(xlRowNum++);
        xlRow.setHeightInPoints(29.25f);
        Cell excelCell = xlRow.createCell(0);
        excelCell.setCellValue(config.getYear());
        excelCell.setCellStyle(cellStyleFactory.yearHeading());
        sheet.setColumnWidth(0, 4064);  // column width 13.86

        for (int i = 1; i <= 31; i++) {
            Cell c = xlRow.createCell(i);
            c.setCellValue(i);
            c.setCellStyle(cellStyleFactory.columnHeading(i));
            sheet.setColumnWidth(i, 1340);  // column width 4.57
        }
    }



    private void addMonthHeadingCell(final Row pRow, final Position pPos)
    {
        final Cell excelCell = pRow.createCell(0);

        if (pPos.isFirstRowOfDay()) {
            excelCell.setCellValue(" " + pPos.getMonth().getDisplayName(TextStyle.FULL, config.getLocale()));
        }
        excelCell.setCellStyle(cellStyleFactory.monthHeading(pPos));
    }



    private void addDayCell(final Row pRow, final Position pPos, final SortedSet<Art> pCategories)
    {
        final Cell excelCell = pRow.createCell(pPos.getDay());

        final boolean isDayEmpty = pCategories == null || pCategories.isEmpty();
        if (!pPos.dayExists()) {
            excelCell.setBlank();
            excelCell.setCellStyle(cellStyleFactory.emptyDay(pPos));
        }
        else if (isDayEmpty) {
            if (pPos.isFirstRowOfDay()) {
                excelCell.setCellValue(getDayOfWeekDisplay(pPos.getDayOfWeek(), config.getLocale()));
            }
            if (pPos.isSunday()) {
                excelCell.setCellStyle(cellStyleFactory.sunday(pPos));
            }
            else {
                excelCell.setCellStyle(cellStyleFactory.emptyDay(pPos));
            }
        }
        else if (pPos.isFirstRowOfDay()) {
            excelCell.setCellValue(getDayOfWeekDisplay(pPos.getDayOfWeek(), config.getLocale()));
            excelCell.setCellStyle(cellStyleFactory.dayHeading(pPos, pCategories.first()));
        }
        else {
            if (pCategories.size() >= pPos.getDayRowIdx()) {
                final Art category = pPos.getDayRowIdx() == 1 ? pCategories.first() : pCategories.last();
                excelCell.setCellValue(category.getKuerzel());
                excelCell.setCellStyle(cellStyleFactory.dayCell(pPos, category));
            }
            else {
                excelCell.setBlank();
                excelCell.setCellStyle(cellStyleFactory.centered(pPos));
            }
            if (pPos.getDayRowIdx() == 2 && pCategories.size() == 1) {
                sheet.addMergedRegion(new CellRangeAddress(xlRowNum - 2, xlRowNum - 1, pPos.getDay(), pPos.getDay()));
            }
        }
    }



    private String getDayOfWeekDisplay(final DayOfWeek pDayOfWeek, final Locale pLocale)
    {
        return pDayOfWeek.getDisplayName(TextStyle.SHORT_STANDALONE, pLocale);
    }



    private void mergeMonthNames()
    {
        for (int month = Calendar.JANUARY; month <= Calendar.DECEMBER; month++) {
            int xlRowNum = month * ROWS_PER_DAY + 2;
            sheet.addMergedRegion(new CellRangeAddress(xlRowNum, xlRowNum + ROWS_PER_DAY - 1, 0, 0));
        }
    }



    private void addNotices()
    {
        sheet.createRow(xlRowNum++);
        Row xlRow = sheet.createRow(xlRowNum++);
        Cell excelCell = xlRow.createCell(1);
        excelCell.setCellValue("Öffnungszeiten Hafen:  Mo. - Fr.:  7 - 12 und 13 - 17 Uhr,   "
            + "Samstag:  8 - 14 Uhr,   Mo./Di. kein Sondermüll");
        excelCell.setCellStyle(cellStyleFactory.noteBig());

        xlRow = sheet.createRow(xlRowNum++);
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

        sheet.createRow(xlRowNum++);  // for inclusion in print area, and extra notes
        sheet.createRow(xlRowNum);
    }



    private void addQrCode()
    {
        int pictureIndex;
        try (final InputStream fis = getClass().getResourceAsStream("qrcode.png")) {
            pictureIndex = workbook.addPicture(fis, XSSFWorkbook.PICTURE_TYPE_PNG);
        }
        catch (IOException e) {
            throw new RuntimeException(e.getMessage(), e);
        }

        final int logoSizeRows = 3;
        final int logoRow = xlRowNum + 1 - logoSizeRows;
        final XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 30, logoRow, 32, logoRow + logoSizeRows);
        anchor.setAnchorType(XSSFClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);

        final XSSFDrawing drawing = sheet.createDrawingPatriarch();
        final XSSFPicture picture = drawing.createPicture(anchor, pictureIndex);
        picture.resize(0.88d, 1);
    }



    private void setPrintSetup()
    {
        final XSSFPrintSetup printSetup = sheet.getPrintSetup();
        workbook.setPrintArea(
            0,   // sheet index
            0,   // start column
            31,  // end column
            0,   // start row
            xlRowNum // end row
        );
        printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
        printSetup.setLandscape(true);
    }



    private void setWorkbookProperties()
    {
        final POIXMLProperties.CoreProperties workbookProps = workbook.getProperties().getCoreProperties();
        workbookProps.setTitle("Abfallkalender " + config.getYear());
        workbookProps.setCreator("abfall");
        workbookProps.setDescription("Generated by \"abfall\" from https://github.com/tsjensen/abfall");
    }
}
