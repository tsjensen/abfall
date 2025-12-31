/*
 * abfall - convert ICS format trash calendar into a single Excel sheet
 * Copyright (C) 2011-2026 Thomas Jensen
 *
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License, version 3, as published by the Free Software Foundation.
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <https://www.gnu.org/licenses/>.
 *
 * SPDX-License-Identifier: GPL-3.0-only
 */
package com.thomasjensen.abfall;

import java.awt.Color;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STGradientType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;


/**
 * Some POI cell styles.
 */
public class CellStyleFactory
{
    private static final IndexedColorMap DEFAULT_COLOR_MAP = new DefaultIndexedColorMap();

    private static final XSSFColor VERTICAL_SEP_COLOR = new XSSFColor(new Color(128, 128, 128), DEFAULT_COLOR_MAP);

    private final Holidays holidays = new Holidays();

    private final XSSFWorkbook workbook;



    public CellStyleFactory(final XSSFWorkbook pWorkbook)
    {
        workbook = pWorkbook;
    }



    public XSSFCellStyle monthHeading(final Position pPos)
    {
        final XSSFCellStyle result = workbook.createCellStyle();
        result.setAlignment(HorizontalAlignment.LEFT);
        result.setVerticalAlignment(VerticalAlignment.CENTER);
        result.setFont(boldCalibri(14));
        if (pPos.isOddMonth()) {
            setOddBackground(result);
        }
        if (pPos.isFirstRowOfDay()) {
            result.setBorderTop(pPos.isJanuary() ? BorderStyle.MEDIUM : BorderStyle.THIN);
        }
        if (pPos.isLastRowOfDay()) {
            result.setBorderBottom(pPos.isDecember() ? BorderStyle.MEDIUM : BorderStyle.THIN);
        }
        result.setBorderRight(BorderStyle.MEDIUM);
        result.setBorderLeft(BorderStyle.MEDIUM);
        return result;
    }



    public XSSFCellStyle yearHeading()
    {
        final XSSFCellStyle result = alignCenter(workbook.createCellStyle());
        result.setFont(boldCalibri(20));
        result.setBorderBottom(BorderStyle.MEDIUM);
        return result;
    }



    public XSSFCellStyle dayHeading(final Position pPos, final Art pCategory)
    {
        final XSSFCellStyle result = alignCenter(workbook.createCellStyle());

        result.setFillPattern(pCategory == Art.Gartenabfall
            ? FillPatternType.THIN_FORWARD_DIAG : FillPatternType.SOLID_FOREGROUND);
        result.setFillForegroundColor(new XSSFColor(pCategory.getColor(), DEFAULT_COLOR_MAP));

        result.setBorderTop(pPos.isJanuary() ? BorderStyle.MEDIUM : BorderStyle.THIN);
        if (pPos.isDay1()) {
            result.setBorderLeft(BorderStyle.MEDIUM);
        }
        else {
            result.setBorderLeft(BorderStyle.THIN);
            result.setLeftBorderColor(VERTICAL_SEP_COLOR);
        }
        if (pPos.isDay31()) {
            result.setBorderRight(BorderStyle.MEDIUM);
        }

        if (holidays.isHoliday(pPos)) {
            final XSSFFont font = workbook.createFont();
            font.setUnderline(FontUnderline.SINGLE);
            result.setFont(font);
        }
        return result;
    }



    public XSSFCellStyle dayCell(final Position pPos, final Art pCategory)
    {
        final XSSFCellStyle result = alignCenter(workbook.createCellStyle());
        if (pPos.isDay1()) {
            result.setBorderLeft(BorderStyle.MEDIUM);
        }
        else {
            result.setBorderLeft(BorderStyle.THIN);
            result.setLeftBorderColor(VERTICAL_SEP_COLOR);
        }
        if (pPos.isDay31()) {
            result.setBorderRight(BorderStyle.MEDIUM);
        }
        if (pPos.isLastRowOfDay()) {
            result.setBorderBottom(pPos.isDecember() ? BorderStyle.MEDIUM : BorderStyle.THIN);
        }

        result.setFillBackgroundColor(new XSSFColor(Color.WHITE, DEFAULT_COLOR_MAP));
        result.setFillForegroundColor(new XSSFColor(Art.Gartenabfall.getColor(), DEFAULT_COLOR_MAP));
        result.setFillPattern(FillPatternType.THIN_FORWARD_DIAG);
        final CTFill ctFill = workbook.getStylesSource().getFillAt((int) result.getCoreXf().getFillId()).getCTFill();
        ctFill.unsetPatternFill();

        if (pCategory == Art.Gartenabfall) {
            final CTPatternFill ctPatternFill = ctFill.addNewPatternFill();
            ctPatternFill.addNewFgColor().setRgb(toBytes(pCategory.getColor()));
            ctPatternFill.setPatternType(STPatternType.LIGHT_UP);
            return result;
        }

        byte[] rgbInner;
        switch (pCategory) {
            case Bio:
                rgbInner = toBytes(new Color(235, 241, 222));
                break;
            case Papier:
                rgbInner = toBytes(new Color(221, 231, 242));
                break;
            case GelberSack:
                rgbInner = toBytes(new Color(255, 255, 204));
                break;
            default: // Rest || Schadstoffmobil
                rgbInner = toBytes(new Color(255, 255, 255));
                break;
        }
        final byte[] rgbOuter = toBytes(pCategory.getColor());

        final CTGradientFill ctGradientFill = ctFill.addNewGradientFill();
        ctGradientFill.setType(STGradientType.PATH);
        ctGradientFill.setLeft(0.5d);
        ctGradientFill.setRight(0.5d);
        ctGradientFill.setTop(0.5d);
        ctGradientFill.setBottom(0.5d);

        ctGradientFill.addNewStop().setPosition(0.0);
        ctGradientFill.getStopArray(0).addNewColor().setRgb(rgbInner);
        ctGradientFill.addNewStop().setPosition(1.0);
        ctGradientFill.getStopArray(1).addNewColor().setRgb(rgbOuter);

        return result;
    }



    public XSSFCellStyle heading()
    {
        final XSSFCellStyle headingStyle = workbook.createCellStyle();
        headingStyle.setFont(boldCalibri(24));
        return headingStyle;
    }



    public XSSFCellStyle columnHeading(final int pDay)
    {
        final XSSFCellStyle result = alignCenter(workbook.createCellStyle());
        result.setFont(boldCalibri(14));
        result.setBorderTop(BorderStyle.MEDIUM);
        result.setBorderBottom(BorderStyle.MEDIUM);
        if (pDay == 1) {
            result.setBorderLeft(BorderStyle.MEDIUM);
        }
        else {
            result.setBorderLeft(BorderStyle.THIN);
            result.setLeftBorderColor(VERTICAL_SEP_COLOR);
        }
        if (pDay == 31) {
            result.setBorderRight(BorderStyle.MEDIUM);
        }
        return result;
    }



    public CellStyle centered(final Position pPos)
    {
        final XSSFCellStyle result = alignCenter(workbook.createCellStyle());
        addNormalCellBorders(pPos, result);
        return result;
    }



    public XSSFCellStyle emptyDay(final Position pPos)
    {
        XSSFCellStyle result = alignCenter(workbook.createCellStyle());
        if (pPos.isOddMonth()) {
            setOddBackground(result);
        }
        addNormalCellBorders(pPos, result);
        return result;
    }



    private void addNormalCellBorders(final Position pPos, final XSSFCellStyle pStyle)
    {
        if (pPos.isFirstRowOfDay()) {
            pStyle.setBorderTop(pPos.isJanuary() ? BorderStyle.MEDIUM : BorderStyle.THIN);
        }
        if (pPos.isDay1()) {
            pStyle.setBorderLeft(BorderStyle.MEDIUM);
        }
        else {
            pStyle.setBorderLeft(BorderStyle.THIN);
            pStyle.setLeftBorderColor(VERTICAL_SEP_COLOR);
        }
        if (pPos.isDay31()) {
            pStyle.setBorderRight(BorderStyle.MEDIUM);
        }
        if (pPos.isLastRowOfDay()) {
            pStyle.setBorderBottom(pPos.isDecember() ? BorderStyle.MEDIUM : BorderStyle.THIN);
        }
    }



    public XSSFCellStyle sunday(final Position pPos)
    {
        final XSSFCellStyle result = alignCenter(workbook.createCellStyle());
        result.setFillForegroundColor(new XSSFColor(new Color(191, 191, 191), DEFAULT_COLOR_MAP));
        result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font = workbook.createFont();
        font.setColor(new XSSFColor(new Color(255, 255, 255), DEFAULT_COLOR_MAP));
        result.setFont(font);
        addNormalCellBorders(pPos, result);
        return result;
    }



    private byte[] toBytes(final Color pColor)
    {
        return new byte[]{
            (byte) pColor.getRed(),
            (byte) pColor.getGreen(),
            (byte) pColor.getBlue()
        };
    }



    private XSSFCellStyle alignCenter(final XSSFCellStyle pCellStyle)
    {
        pCellStyle.setAlignment(HorizontalAlignment.CENTER);
        pCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return pCellStyle;
    }



    private void setOddBackground(final XSSFCellStyle pCellStyle)
    {
        pCellStyle.setFillForegroundColor(new XSSFColor(new Color(221, 221, 221), DEFAULT_COLOR_MAP));
        pCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }



    private XSSFFont boldCalibri(final int pFontSizePx)
    {
        final XSSFFont result = workbook.createFont();
        result.setFontName("Calibri");
        result.setFontHeightInPoints((short) pFontSizePx);
        result.setBold(true);
        return result;
    }



    public XSSFCellStyle noteBig()
    {
        final XSSFCellStyle result = workbook.createCellStyle();
        XSSFFont font = boldCalibri(14);
        font.setBold(false);
        result.setFont(font);
        return result;
    }



    public CellStyle noteSmall()
    {
        final XSSFCellStyle result = workbook.createCellStyle();
        XSSFFont font = boldCalibri(11);
        font.setBold(false);
        result.setFont(font);
        return result;
    }
}
