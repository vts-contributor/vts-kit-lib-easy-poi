package com.atviettelsolutions.easypoi.poi.excel;

import com.atviettelsolutions.easypoi.poi.excel.html.ExcelToHtmlServer;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel into an interface
 *
 * @author caprocute
 */
public final class ExcelToHtmlUtil {

    private ExcelToHtmlUtil() {
    }

    /**
     * Convert to Table
     *
     * @param wb Excel
     * @return
     */
    public static String toTableHtml(Workbook wb) {
        return new ExcelToHtmlServer(wb, false, 0).printPage();
    }

    /**
     * Convert to Table
     *
     * @param wb       Excel
     * @param sheetNum sheetNum
     * @return
     */
    public static String toTableHtml(Workbook wb, int sheetNum) {
        return new ExcelToHtmlServer(wb, false, sheetNum).printPage();
    }

    /**
     * CONVERT TO A FULL INTERFACE
     *
     * @param wb       Excel
     * @return
     */
    public static String toAllHtml(Workbook wb) {
        return new ExcelToHtmlServer(wb, true, 0).printPage();
    }

    /**
     * Convert to a full interface
     *
     * @param wb       Excel
     * @param sheetNum sheetNum
     * @return
     */
    public static String toAllHtml(Workbook wb, int sheetNum) {
        return new ExcelToHtmlServer(wb, true, sheetNum).printPage();
    }

}
