
package com.viettel.vtskit.easypoi.poi.excel.export;

import com.viettel.vtskit.easypoi.poi.common.PoiExcelGraphDataUtil;
import com.viettel.vtskit.easypoi.poi.common.PoiPublicUtil;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelTarget;
import com.viettel.vtskit.easypoi.poi.excel.entity.ExportParams;
import com.viettel.vtskit.easypoi.poi.excel.entity.enmus.ExcelType;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.viettel.vtskit.easypoi.poi.excel.entity.vo.PoiBaseConstants;
import com.viettel.vtskit.easypoi.poi.excel.export.base.ExcelExportBase;
import com.viettel.vtskit.easypoi.poi.excel.export.styler.IExcelExportStyler;
import com.viettel.vtskit.easypoi.poi.exeption.ExcelExportException;
import com.viettel.vtskit.easypoi.poi.exeption.excel.enums.ExcelExportEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Excel export service
 *
 * @author caprocute
 */
public class ExcelExportServer extends ExcelExportBase {

    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelExportServer.class);

    // The maximum number of rows, which exceeds the automatic multi-sheet
    private int MAX_NUM = 60000;

    protected int createHeaderAndTitle(ExportParams entity, Sheet sheet, Workbook workbook, List<ExcelExportEntity> excelParams) {
        int rows = 0, feildWidth = getFieldWidth(excelParams);
        if (entity.getTitle() != null) {
            rows += createHeaderRow(entity, sheet, workbook, feildWidth);
        }
        rows += createTitleRow(entity, sheet, workbook, rows, excelParams);
        if (entity.isFixedTitle()) {
            sheet.createFreezePane(0, rows, 0, rows);
        }
        return rows;
    }

    /**
     * Create a table header change
     *
     * @param entity
     * @param sheet
     * @param workbook
     * @param feildWidth
     */
    public int createHeaderRow(ExportParams entity, Sheet sheet, Workbook workbook, int feildWidth) {
        Row row = sheet.createRow(0);
        row.setHeight(entity.getTitleHeight());
        createStringCell(row, 0, entity.getTitle(), getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
        for (int i = 1; i <= feildWidth; i++) {
            createStringCell(row, i, "", getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
        }
        try {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, feildWidth));
        } catch (IllegalArgumentException e) {
            LOGGER.error("Merge cell error logs：" + e.getMessage());
            e.fillInStackTrace();
        }
        if (entity.getSecondTitle() != null) {
            row = sheet.createRow(1);
            row.setHeight(entity.getSecondTitleHeight());
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.RIGHT);
            createStringCell(row, 0, entity.getSecondTitle(), style, null);
            for (int i = 1; i <= feildWidth; i++) {
                createStringCell(row, i, "", getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
            }
            try {
                sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, feildWidth));
            } catch (IllegalArgumentException e) {
                LOGGER.error("Merge cell error logs：" + e.getMessage());
                e.fillInStackTrace();
            }
            return 2;
        }
        return 1;
    }

    public void createSheet(Workbook workbook, ExportParams entity, Class<?> pojoClass, Collection<?> dataSet, String[] exportFields) {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel export start ,class is {}", pojoClass);
            LOGGER.debug("Excel version is {}", entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || pojoClass == null || dataSet == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        super.type = entity.getType();
        if (type.equals(ExcelType.XSSF)) {
            MAX_NUM = 1000000;
        }
        Sheet sheet = null;
        try {
            sheet = workbook.createSheet(entity.getSheetName());
        } catch (Exception e) {
            // Repeated traversal, duplicate names occur, and a non-specified name Sheet is created
            sheet = workbook.createSheet();
        }
        try {
            dataHanlder = entity.getDataHanlder();
            if (dataHanlder != null) {
                String[] needHandlerFields = dataHanlder.getNeedHandlerFields();
                if (needHandlerFields != null && needHandlerFields.length > 0) {
                    needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
                }
            }
            // Create a table style
            setExcelExportStyler((IExcelExportStyler) entity.getStyle().getConstructor(Workbook.class).newInstance(workbook));
            Drawing patriarch = sheet.createDrawingPatriarch();
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            // Get all fields
            Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);

            //Custom export fields are supported
            if (exportFields != null) {
                List<Field> list = new ArrayList<Field>(Arrays.asList(fileds));
                for (int i = 0; i < list.size(); i++) {
                    if (!Arrays.asList(exportFields).contains(list.get(i).getName())) {
                        list.remove(i);
                        i--;
                    }
                }

                if (list != null && list.size() > 0) {
                    fileds = list.toArray(new Field[0]);
                } else {
                    fileds = null;
                }
            }
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            String targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass, null);
            reConfigExcelExportParams(excelParams, entity);
            int index = entity.isCreateHeadRows() ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
            int titleHeight = index;
            setCellWith(excelParams, sheet);
            setColumnHidden(excelParams, sheet);

            short rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) {
                Object t = its.next();
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
                tempList.add(t);
                if (index >= MAX_NUM)
                    break;
            }
            mergeCells(sheet, excelParams, titleHeight);

            if (entity.getFreezeCol() != 0) {
                sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
            }

            its = dataSet.iterator();
            for (int i = 0, le = tempList.size(); i < le; i++) {
                its.next();
                its.remove();
            }
            // Create totals
            addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);

            // It was found that there were still lists left, and continued to create the sheet in a loop
            if (dataSet.size() > 0) {
                createSheet(workbook, entity, pojoClass, dataSet, exportFields);
            }

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    public void createSheetForMap(Workbook workbook, ExportParams entity, List<ExcelExportEntity> entityList, Collection<? extends Map<?, ?>> dataSet) {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel version is {}", entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || entityList == null || dataSet == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        super.type = entity.getType();
        if (type.equals(ExcelType.XSSF)) {
            MAX_NUM = 1000000;
        }
        Sheet sheet = null;
        try {
            sheet = workbook.createSheet(entity.getSheetName());
        } catch (Exception e) {
            // Repeated traversal, duplicate names occur, and a non-specified name Sheet is created
            sheet = workbook.createSheet();
        }
        try {
            dataHanlder = entity.getDataHanlder();
            if (dataHanlder != null) {
                needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
            }
            // Create a table style
            setExcelExportStyler((IExcelExportStyler) entity.getStyle().getConstructor(Workbook.class).newInstance(workbook));
            Drawing patriarch = sheet.createDrawingPatriarch();
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            int index = entity.isCreateHeadRows() ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
            int titleHeight = index;
            setCellWith(excelParams, sheet);
            setColumnHidden(excelParams, sheet);
            short rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) {
                Object t = its.next();
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
                tempList.add(t);
                if (index >= MAX_NUM)
                    break;
            }
            if (entity.getFreezeCol() != 0) {
                sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
            }

            mergeCells(sheet, excelParams, titleHeight);

            its = dataSet.iterator();
            for (int i = 0, le = tempList.size(); i < le; i++) {
                its.next();
                its.remove();
            }
            // It was found that there were still lists left, and continued to create the sheet in a loop
            if (dataSet.size() > 0) {
                createSheetForMap(workbook, entity, entityList, dataSet);
            }

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    /**
     * Create a table header
     *
     * @param title
     * @param index
     */
    private int createTitleRow(ExportParams title, Sheet sheet, Workbook workbook, int index, List<ExcelExportEntity> excelParams) {
        Row row = sheet.createRow(index);
        int rows = getRowNums(excelParams);
        row.setHeight((short) 450);
        Row listRow = null;
        if (rows == 2) {
            listRow = sheet.createRow(index + 1);
            listRow.setHeight((short) 450);
        }
        int cellIndex = 0;
        CellStyle titleStyle = getExcelExportStyler().getTitleStyle(title.getColor());
        for (int i = 0, exportFieldTitleSize = excelParams.size(); i < exportFieldTitleSize; i++) {
            ExcelExportEntity entity = excelParams.get(i);
            if (entity.isColspan()) {
                List<String> subList = entity.getSubColumnList();
                if (subList == null || subList.size() == 0) {
                    continue;
                } else {
                    entity.initSubExportEntity(excelParams);
                }
            }
            if (StringUtils.isNotBlank(entity.getName())) {
                createStringCell(row, cellIndex, entity.getName(), titleStyle, entity);
            }
            if (entity.getList() != null) {
                List<ExcelExportEntity> sTitel = entity.getList();
                if (StringUtils.isNotBlank(entity.getName())) {
                    try {
                        sheet.addMergedRegion(new CellRangeAddress(index, index, cellIndex, cellIndex + sTitel.size() - 1));
                    } catch (IllegalArgumentException e) {
                        LOGGER.error("Merge cell error logs：" + e.getMessage());
                        e.fillInStackTrace();
                    }
                }
                for (int j = 0, size = sTitel.size(); j < size; j++) {
                    createStringCell(rows == 2 ? listRow : row, cellIndex, sTitel.get(j).getName(), titleStyle, entity);
                    cellIndex++;
                }
                cellIndex--;
            } else if (rows == 2) {
                createStringCell(listRow, cellIndex, "", titleStyle, entity);
                try {
                    sheet.addMergedRegion(new CellRangeAddress(index, index + 1, cellIndex, cellIndex));
                } catch (IllegalArgumentException e) {
                    LOGGER.error("Merge cell error logs：" + e.getMessage());
                    e.fillInStackTrace();
                }
            }
            cellIndex++;
        }
        return rows;

    }

    /**
     * Determine whether the header has only one or two rows
     *
     * @param excelParams
     * @return
     */
    private int getRowNums(List<ExcelExportEntity> excelParams) {
        for (int i = 0; i < excelParams.size(); i++) {
            ExcelExportEntity temp = excelParams.get(i);
            if ((temp.getList() != null || temp.isColspan()) && StringUtils.isNotBlank(temp.getName())) {
                return 2;
            }
        }
        return 1;
    }

    protected ExcelExportEntity indexExcelEntity(ExportParams entity) {
        ExcelExportEntity exportEntity = new ExcelExportEntity();
        exportEntity.setOrderNum(0);
        exportEntity.setName(entity.getIndexName());
        exportEntity.setWidth(10);
        exportEntity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
        return exportEntity;
    }

    /**
     * Add data to the sheet
     *
     * @param workbook
     * @param entity
     * @param entityList
     * @param dataSet
     * @param sheet
     */
    protected void insertDataToSheet(Workbook workbook, ExportParams entity,
                                     List<ExcelExportEntity> entityList, Collection<? extends Map<?, ?>> dataSet,
                                     Sheet sheet) {
        try {
            dataHanlder = entity.getDataHanlder();
            if (dataHanlder != null && dataHanlder.getNeedHandlerFields() != null) {
                needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
            }
            // Create a table style
            setExcelExportStyler((IExcelExportStyler) entity.getStyle()
                    .getConstructor(Workbook.class).newInstance(workbook));
            Drawing patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            int index = entity.isCreateHeadRows()
                    ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
            int titleHeight = index;
            setCellWith(excelParams, sheet);
            setColumnHidden(excelParams, sheet);
            short rowHeight = entity.getHeight() != 0 ? entity.getHeight() : getRowHeight(excelParams);
            setCurrentIndex(1);
            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) {
                Object t = its.next();
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight, 0)[0];
                tempList.add(t);
                if (index >= MAX_NUM) {
                    break;
                }
            }
            if (entity.getFreezeCol() != 0) {
                sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
            }

            mergeCells(sheet, excelParams, titleHeight);

            its = dataSet.iterator();
            for (int i = 0, le = tempList.size(); i < le; i++) {
                its.next();
                its.remove();
            }
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug("List data more than max ,data size is {}",
                        dataSet.size());
            }
            // It was found that there were still lists left, and continued to create the sheet in a loop
            if (dataSet.size() > 0) {
                createSheetForMap(workbook, entity, entityList, dataSet);
            } else {
                // Create totals
                addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);
            }

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }
}
