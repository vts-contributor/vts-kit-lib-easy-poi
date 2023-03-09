package com.viettel.vtskit.easypoi.poi.common;

import com.viettel.vtskit.easypoi.poi.excel.entity.params.MergeEntity;
import com.viettel.vtskit.easypoi.poi.exeption.ExcelExportException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

/**
 * Merge cells vertically
 *
 * @author caprocute
 */
public final class PoiMergeCellUtil {
    private final static Logger LOGGER = LoggerFactory.getLogger(PoiMergeCellUtil.class);

    private PoiMergeCellUtil() {
    }

    /**
     * Vertically merge cells with the same content
     *
     * @param sheet
     * @param startRow
     * @param columns
     */
    public static void mergeCells(Sheet sheet, int startRow, Integer... columns) {
        if (columns == null) {
            throw new ExcelExportException("At least 1 column needs to be processed");
        }
        Map<Integer, int[]> mergeMap = new HashMap<Integer, int[]>();
        for (int i = 0; i < columns.length; i++) {
            mergeMap.put(columns[i], null);
        }
        mergeCells(sheet, mergeMap, startRow, sheet.getLastRowNum());
    }

    /**
     * Vertically merge cells with the same content
     *
     * @param sheet
     * @param mergeMap key--column, value--dependent column, no empty
     * @param startRow start-line
     */
    public static void mergeCells(Sheet sheet, Map<Integer, int[]> mergeMap, int startRow) {
        mergeCells(sheet, mergeMap, startRow, sheet.getLastRowNum());
    }

    /**
     * Vertically merge cells with the same content
     *
     * @param sheet
     * @param mergeMap --column, value--dependent column, no empty
     * @param startRow
     * @param endRow
     */
    public static void mergeCells(Sheet sheet, Map<Integer, int[]> mergeMap, int startRow, int endRow) {
        Map<Integer, MergeEntity> mergeDataMap = new HashMap<Integer, MergeEntity>();
        if (mergeMap.size() == 0) {
            return;
        }
        Row row;
        Set<Integer> sets = mergeMap.keySet();
        String text;
        for (int i = startRow; i <= endRow; i++) {
            row = sheet.getRow(i);
            for (Integer index : sets) {
                if (row == null || row.getCell(index) == null) {
                    if (mergeDataMap.get(index) == null) {
                        continue;
                    }
                    if (mergeDataMap.get(index).getEndRow() == 0) {
                        mergeDataMap.get(index).setEndRow(i - 1);
                    }
                } else {
                    text = row.getCell(index).getStringCellValue();
                    if (StringUtils.isNotEmpty(text)) {
                        hanlderMergeCells(index, i, text, mergeDataMap, sheet, row.getCell(index), mergeMap.get(index));
                    } else {
                        mergeCellOrContinue(index, mergeDataMap, sheet);
                    }
                }
            }
        }
        if (mergeDataMap.size() > 0) {
            for (Integer index : mergeDataMap.keySet()) {
                try {
                    sheet.addMergedRegion(new CellRangeAddress(mergeDataMap.get(index).getStartRow(), mergeDataMap.get(index).getEndRow(), index, index));
                } catch (IllegalArgumentException e) {
                    LOGGER.error("Merge cell error log：" + e.getMessage());
                    e.fillInStackTrace();
                }
            }
        }

    }

    /**
     * Handle merged cells     *
     * @param index
     * @param rowNum
     * @param text
     * @param mergeDataMap
     * @param sheet
     * @param cell
     * @param delys
     */
    private static void hanlderMergeCells(Integer index, int rowNum, String text, Map<Integer, MergeEntity> mergeDataMap, Sheet sheet, Cell cell, int[] delys) {
        if (mergeDataMap.containsKey(index)) {
            if (checkIsEqualByCellContents(mergeDataMap.get(index), text, cell, delys, rowNum)) {
                mergeDataMap.get(index).setEndRow(rowNum);
            } else {
                try {
                    sheet.addMergedRegion(new CellRangeAddress(mergeDataMap.get(index).getStartRow(), mergeDataMap.get(index).getEndRow(), index, index));
                } catch (IllegalArgumentException e) {
                    LOGGER.error("Merge cell error log：" + e.getMessage());
                    e.fillInStackTrace();
                }
                mergeDataMap.put(index, createMergeEntity(text, rowNum, cell, delys));
            }
        } else {
            mergeDataMap.put(index, createMergeEntity(text, rowNum, cell, delys));
        }
    }

    /**
     * Judging when the character is empty
     *
     * @param index
     * @param mergeDataMap
     * @param sheet
     */
    private static void mergeCellOrContinue(Integer index, Map<Integer, MergeEntity> mergeDataMap, Sheet sheet) {
        if (mergeDataMap.containsKey(index) && mergeDataMap.get(index).getEndRow() != mergeDataMap.get(index).getStartRow()) {
            try {
                sheet.addMergedRegion(new CellRangeAddress(mergeDataMap.get(index).getStartRow(), mergeDataMap.get(index).getEndRow(), index, index));
            } catch (IllegalArgumentException e) {
                LOGGER.error("Merge cell error log：" + e.getMessage());
                e.fillInStackTrace();
            }
            mergeDataMap.remove(index);
        }
    }

    private static MergeEntity createMergeEntity(String text, int rowNum, Cell cell, int[] delys) {
        MergeEntity mergeEntity = new MergeEntity(text, rowNum, rowNum);
        // There is a dependency
        if (delys != null && delys.length != 0) {
            List<String> list = new ArrayList<String>(delys.length);
            mergeEntity.setRelyList(list);
            for (int i = 0; i < delys.length; i++) {
                list.add(getCellNotNullText(cell, delys[i], rowNum));
            }
        }
        return mergeEntity;
    }

    private static boolean checkIsEqualByCellContents(MergeEntity mergeEntity, String text, Cell cell, int[] delys, int rowNum) {
        // no dependencies
        if (delys == null || delys.length == 0) {
            return mergeEntity.getText().equals(text);
        }
        // There is a dependency
        if (mergeEntity.getText().equals(text)) {
            for (int i = 0; i < delys.length; i++) {
                if (!getCellNotNullText(cell, delys[i], rowNum).equals(mergeEntity.getRelyList().get(i))) {
                    return false;
                }
            }
            return true;
        }
        return false;
    }

    /**
     * Get the value of a cell, make sure that the cell must have a value, otherwise query up
     *
     * @param cell
     * @param index
     * @param rowNum
     * @return
     */
    private static String getCellNotNullText(Cell cell, int index, int rowNum) {
        String temp = cell.getRow().getCell(index).getStringCellValue();
        while (StringUtils.isEmpty(temp)) {
            temp = cell.getRow().getSheet().getRow(--rowNum).getCell(index).getStringCellValue();
        }
        return temp;
    }

    /**
     * process merge
     * @param sheet
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public static void addMergedRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        try {
            //Add merged range of cells
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        } catch (Exception e) {
            LOGGER.debug("A merge cell error has occurred,{},{},{},{}", new Integer[]{
                    firstRow, lastRow, firstCol, lastCol
            });
            // Ignore merged errors, do not print exceptions
            LOGGER.debug(e.getMessage(), e);
        }
    }

}
