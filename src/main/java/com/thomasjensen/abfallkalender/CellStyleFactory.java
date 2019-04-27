package com.thomasjensen.abfallkalender;

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
            withOddBackground(result);
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
        switch (pCategory) {
            case Bio:
                result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                result.setFillForegroundColor(new XSSFColor(new Color(196, 215, 155), DEFAULT_COLOR_MAP));
                break;
            case Papier:
                result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                result.setFillForegroundColor(new XSSFColor(new Color(149, 179, 215), DEFAULT_COLOR_MAP));
                break;
            case GelberSack:
                result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                result.setFillForegroundColor(new XSSFColor(new Color(255, 255, 102), DEFAULT_COLOR_MAP));
                break;
            case Rest:
                result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                result.setFillForegroundColor(new XSSFColor(new Color(207, 121, 119), DEFAULT_COLOR_MAP));
                break;
            case Gartenabfall:
                result.setFillPattern(FillPatternType.THIN_FORWARD_DIAG);
                result.setFillForegroundColor(new XSSFColor(new Color(0, 176, 80), DEFAULT_COLOR_MAP));
                break;
            default: // Schadstoffmobil
                result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                result.setFillForegroundColor(new XSSFColor(new Color(250, 192, 146), DEFAULT_COLOR_MAP));
                break;
        }

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
        result.setFillForegroundColor(new XSSFColor(new Color(0, 176, 80), DEFAULT_COLOR_MAP));
        result.setFillPattern(FillPatternType.THIN_FORWARD_DIAG);
        final CTFill ctFill = workbook.getStylesSource().getFillAt((int) result.getCoreXf().getFillId()).getCTFill();
        ctFill.unsetPatternFill();

        if (pCategory == Art.Gartenabfall) {
            final CTPatternFill ctPatternFill = ctFill.addNewPatternFill();
            ctPatternFill.addNewFgColor().setRgb(toBytes(new Color(0, 176, 80)));
            ctPatternFill.setPatternType(STPatternType.LIGHT_UP);
            return result;
        }

        byte[] rgbInner;
        byte[] rgbOuter;
        switch (pCategory) {
            case Bio:
                rgbInner = toBytes(new Color(235, 241, 222));
                rgbOuter = toBytes(new Color(196, 215, 155));
                break;
            case Papier:
                rgbInner = toBytes(new Color(221, 231, 242));
                rgbOuter = toBytes(new Color(150, 180, 216));
                break;
            case GelberSack:
                rgbInner = toBytes(new Color(255, 255, 204));
                rgbOuter = toBytes(new Color(255, 255, 102));
                break;
            case Rest:
                rgbInner = toBytes(new Color(255, 255, 255));
                rgbOuter = toBytes(new Color(207, 121, 119));
                break;
            default: // Schadstoffmobil
                rgbInner = toBytes(new Color(255, 255, 255));
                rgbOuter = toBytes(new Color(250, 192, 146));
                break;
        }

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



    public XSSFCellStyle oddRowCentered(final Position pPos)
    {
        final XSSFCellStyle result = alignCenter(withOddBackground(workbook.createCellStyle()));
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



    private XSSFCellStyle withOddBackground(final XSSFCellStyle pCellStyle)
    {
        pCellStyle.setFillForegroundColor(new XSSFColor(new Color(221, 221, 221), DEFAULT_COLOR_MAP));
        pCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return pCellStyle;
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
