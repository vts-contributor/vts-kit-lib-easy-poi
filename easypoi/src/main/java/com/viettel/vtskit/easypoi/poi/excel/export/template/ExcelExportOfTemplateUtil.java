package com.viettel.vtskit.easypoi.poi.excel.export.template;

import com.viettel.vtskit.easypoi.poi.common.*;
import com.viettel.vtskit.easypoi.poi.exeption.ExcelExportException;
import com.viettel.vtskit.easypoi.poi.exeption.excel.enums.ExcelExportEnum;
import com.viettel.vtskit.easypoi.poi.model.ImageModel;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.viettel.vtskit.easypoi.poi.cache.ExcelCache;
import com.viettel.vtskit.easypoi.poi.cache.ImageCache;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelTarget;
import com.viettel.vtskit.easypoi.poi.excel.entity.TemplateExportParams;
import com.viettel.vtskit.easypoi.poi.excel.entity.enmus.ExcelType;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelForEachParams;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelTemplateParams;
import com.viettel.vtskit.easypoi.poi.excel.export.base.ExcelExportBase;
import com.viettel.vtskit.easypoi.poi.excel.export.styler.IExcelExportStyler;
import com.viettel.vtskit.easypoi.poi.excel.html.helper.MergedRegionHelper;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

import static com.viettel.vtskit.easypoi.poi.common.PoiElUtil.*;


/**
 * Excel Export is exported according to the template
 *
 * @author caprocute
 * @version 1.0
 */
public final class ExcelExportOfTemplateUtil extends ExcelExportBase {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelExportOfTemplateUtil.class);

    /**
     * Cache the cells created by TEMP's for each, skip the template syntax lookup of this cell, and improve efficiency
     */
    private Set<String> tempCreateCellSet = new HashSet<String>();
    /**
     * Template parameters, used globally
     */
    private TemplateExportParams teplateParams;
    /**
     * Cell merge information
     */
    private MergedRegionHelper mergedRegionHelper;


    /**
     * Fill the Sheet with normal data, according to the header information Use the imported partial logic to sit object mapping
     *
     * @param pojoClass
     * @param dataSet
     * @param workbook
     */
    private void addDataToSheet(Class<?> pojoClass, Collection<?> dataSet, Sheet sheet, Workbook workbook) throws Exception {

        if (workbook instanceof XSSFWorkbook) {
            super.type = ExcelType.XSSF;
        }
        // Get the header data
        Map<String, Integer> titlemap = getTitleMap(sheet);
        Drawing patriarch = sheet.createDrawingPatriarch();
        // Get all fields
        Field[] fileds = PoiPublicUtil.getClassFields(pojoClass);
        ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
        String targetId = null;
        if (etarget != null) {
            targetId = etarget.value();
        }
        // Gets the exported data for an entity object
        List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
        getAllExcelField(null, targetId, fileds, excelParams, pojoClass, null);
        // Filter sorting based on table headers
        sortAndFilterExportField(excelParams, titlemap);
        short rowHeight = getRowHeight(excelParams);
        int index = teplateParams.getHeadingRows() + teplateParams.getHeadingStartRow(), titleHeight = index;
        // Move the data down to simulate insertion
        sheet.shiftRows(teplateParams.getHeadingRows() + teplateParams.getHeadingStartRow(),
                sheet.getLastRowNum(), getShiftRows(dataSet, excelParams), true, true);
        if (excelParams.size() == 0) {
            return;
        }
        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
        }
        // Merge homogeneous items
        mergeCells(sheet, excelParams, titleHeight);
    }

    /**
     * Move the data down
     *
     * @param excelParams
     * @return
     */
    private int getShiftRows(Collection<?> dataSet, List<ExcelExportEntity> excelParams) throws Exception {
        int size = 0;
        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            size += getOneObjectSize(t, excelParams);
        }
        return size;
    }

    /**
     * Get the height of a single object, mainly to handle a bunch of more cases
     *
     * @throws Exception
     */
    public int getOneObjectSize(Object t, List<ExcelExportEntity> excelParams) throws Exception {
        ExcelExportEntity entity;
        int maxHeight = 1;
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.getList() != null) {
                Collection<?> list = (Collection<?>) entity.getMethod().invoke(t, new Object[]{});
                if (list != null && list.size() > maxHeight) {
                    maxHeight = list.size();
                }
            }
        }
        return maxHeight;

    }

    public Workbook createExcleByTemplate(TemplateExportParams params, Class<?> pojoClass, Collection<?> dataSet, Map<String, Object> map) {
        // step 1. Determine the address of the template
        if (params == null || map == null || StringUtils.isEmpty(params.getTemplateUrl())) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        Workbook wb = null;
        // step 2. Determine the Excel type of the template and parse the template
        try {
            this.teplateParams = params;
            wb = getCloneWorkBook();
            // Create a table style
            setExcelExportStyler((IExcelExportStyler) teplateParams.getStyle().getConstructor(Workbook.class).newInstance(wb));
            // step 3. Resolve the template
            for (int i = 0, le = params.isScanAllsheet() ? wb.getNumberOfSheets() : params.getSheetNum().length; i < le; i++) {
                if (params.getSheetName() != null && params.getSheetName().length > i && StringUtils.isNotEmpty(params.getSheetName()[i])) {
                    wb.setSheetName(i, params.getSheetName()[i]);
                }
                tempCreateCellSet.clear();
                parseTemplate(wb.getSheetAt(i), map, params.isColForEach());
            }
            if (dataSet != null) {
                // step 4. Normal data filling
                dataHanlder = params.getDataHanlder();
                if (dataHanlder != null) {
                    needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
                }
                addDataToSheet(pojoClass, dataSet, wb.getSheetAt(params.getDataSheetNum()), wb);
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            return null;
        }
        return wb;
    }

    /**
     * Clone Excel prevents manipulation of the original object, and the workbook cannot be cloned, only Excel can be cloned
     *
     * @throws Exception
     * @author caprocute
     * @date 2013-11-11
     */
    private Workbook getCloneWorkBook() throws Exception {
        return ExcelCache.getWorkbookByTemplate(teplateParams.getTemplateUrl(), teplateParams.getSheetNum(), teplateParams.isScanAllsheet());
    }

    /**
     * Obtain the header data and set the sequence number of the header
     *
     * @param sheet
     * @return
     */
    private Map<String, Integer> getTitleMap(Sheet sheet) {
        Row row = null;
        Iterator<Cell> cellTitle;
        Map<String, Integer> titlemap = new HashMap<String, Integer>();
        for (int j = 0; j < teplateParams.getHeadingRows(); j++) {
            row = sheet.getRow(j + teplateParams.getHeadingStartRow());
            cellTitle = row.cellIterator();
            int i = row.getFirstCellNum();
            while (cellTitle.hasNext()) {
                Cell cell = cellTitle.next();
                String value = cell.getStringCellValue();
                if (!StringUtils.isEmpty(value)) {
                    titlemap.put(value, i);
                }
                i = i + 1;
            }
        }
        return titlemap;

    }

    private void parseTemplate(Sheet sheet, Map<String, Object> map, boolean colForeach) throws Exception {
        deleteCell(sheet, map);
        mergedRegionHelper = new MergedRegionHelper(sheet);
        if (colForeach) {
            colForeach(sheet, map);
        }
        Row row = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                if (row.getCell(i) != null && !tempCreateCellSet.contains(row.getRowNum() + "_" + row.getCell(i).getColumnIndex())) {
                    setValueForCellByMap(row.getCell(i), map);
                }
            }
        }
    }

    /**
     * Judge the deletion first, so as not to affect the efficiency
     *
     * @param sheet
     * @param map
     * @throws Exception
     */
    private void deleteCell(Sheet sheet, Map<String, Object> map) throws Exception {
        Row row = null;
        Cell cell = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                cell = row.getCell(i);
                if (row.getCell(i) != null && (cell.getCellTypeEnum() == CellType.STRING || cell.getCellTypeEnum() == CellType.NUMERIC)) {
                    cell.setCellType(CellType.STRING);
                    String text = cell.getStringCellValue();
                    if (text.contains(IF_DELETE)) {
                        if (Boolean.valueOf(eval(text.substring(text.indexOf(START_STR) + 2, text.indexOf(END_STR)).trim(), map).toString())) {
                            PoiSheetUtility.deleteColumn(sheet, i);
                        }
                        cell.setCellValue("");
                    }
                }
            }
        }
    }

    /**
     * Give each cell a parse set value
     *
     * @param cell
     * @param map
     */
    private void setValueForCellByMap(Cell cell, Map<String, Object> map) throws Exception {
        CellType cellType = cell.getCellTypeEnum();
        if (cellType != CellType.STRING && cellType != CellType.NUMERIC) {
            return;
        }
        String oldString;
        cell.setCellType(CellType.STRING);
        oldString = cell.getStringCellValue();
        if (oldString != null && oldString.indexOf(START_STR) != -1 && !oldString.contains(FOREACH)) {
            // step 2. Determine whether there is a parse function
            String params = null;
            boolean isNumber = false;
            if (isNumber(oldString)) {
                isNumber = true;
                oldString = oldString.replace(NUMBER_SYMBOL, "");
            }
            while (oldString.indexOf(START_STR) != -1) {
                params = oldString.substring(oldString.indexOf(START_STR) + 2, oldString.indexOf(END_STR));

                oldString = oldString.replace(START_STR + params + END_STR, eval(params, map).toString());
            }
            // How is the numeric type, set according to the numeric type
            if (isNumber && StringUtils.isNotBlank(oldString)) {
                cell.setCellValue(Double.parseDouble(oldString));
                cell.setCellType(CellType.NUMERIC);
            } else {
                cell.setCellValue(oldString);
            }
        }
        // Judge foreach this method
        if (oldString != null && oldString.contains(FOREACH)) {
            addListDataToExcel(cell, map, oldString.trim());
        }

    }

    private boolean isNumber(String text) {
        return text.startsWith(NUMBER_SYMBOL) || text.contains("{" + NUMBER_SYMBOL) || text.contains(" " + NUMBER_SYMBOL);
    }

    /**
     * Use the foreach loop to output data
     *
     * @param cell
     * @param map
     * @throws Exception
     */
    private void addListDataToExcel(Cell cell, Map<String, Object> map, String name) throws Exception {
        boolean isCreate = !name.contains(FOREACH_NOT_CREATE);
        boolean isShift = name.contains(FOREACH_AND_SHIFT);
        name = name.replace(FOREACH_NOT_CREATE, EMPTY).replace(FOREACH_AND_SHIFT, EMPTY).replace(FOREACH, EMPTY).replace(START_STR, EMPTY);
        String[] keys = name.replaceAll("\\s{1,}", " ").trim().split(" ");
        Collection<?> datas = (Collection<?>) PoiPublicUtil.getParamsValue(keys[0], map);
        Object[] columnsInfo = getAllDataColumns(cell, name.replace(keys[0], EMPTY),
                mergedRegionHelper);
        int rowspan = (Integer) columnsInfo[0], colspan = (Integer) columnsInfo[1];
        List<ExcelForEachParams> columns = (List<ExcelForEachParams>) columnsInfo[2];
        if (datas == null) {
            return;
        }
        Iterator<?> its = datas.iterator();
        Row row;
        int rowIndex = cell.getRow().getRowNum() + 1;
        int loopSize = 0;
        if (its.hasNext()) {
            Object t = its.next();
            cell.getRow().setHeight(columns.get(0).getHeight());
            loopSize = setForeachRowCellValue(isCreate, cell.getRow(), cell.getColumnIndex(), t, columns, map,
                    rowspan, colspan, mergedRegionHelper)[0];
            rowIndex += rowspan - 1 + loopSize - 1;
        }
        //Regardless of whether there is data behind the repair, the insert operation should be performed
        if (isShift && datas.size() * rowspan > 1 && cell.getRowIndex() + rowspan <= cell.getRow().getSheet().getLastRowNum()) {
            int lastRowNum = cell.getRow().getSheet().getLastRowNum();
            int shiftRows = lastRowNum - cell.getRowIndex() - rowspan;
            cell.getRow().getSheet().shiftRows(cell.getRowIndex() + rowspan, lastRowNum, (datas.size() - 1) * rowspan, true, true);
        }
        while (its.hasNext()) {
            Object t = its.next();
            row = createRow(rowIndex, cell.getSheet(), isCreate, rowspan);
            row.setHeight(columns.get(0).getHeight());
            loopSize = setForeachRowCellValue(isCreate, row, cell.getColumnIndex(), t, columns, map, rowspan,
                    colspan, mergedRegionHelper)[0];
            rowIndex += rowspan + loopSize - 1;
        }
    }

    private void setForEeachCellValue(boolean isCreate, Row row, int columnIndex, Object t, List<ExcelTemplateParams> columns, Map<String, Object> map) throws Exception {
        for (int i = 0, max = columnIndex + columns.size(); i < max; i++) {
            if (row.getCell(i) == null)
                row.createCell(i);
        }
        for (int i = 0, max = columns.size(); i < max; i++) {
            boolean isNumber = false;
            String tempStr = new String(columns.get(i).getName());
            if (isNumber(tempStr)) {
                isNumber = true;
                tempStr = tempStr.replace(NUMBER_SYMBOL, "");
            }
            map.put(teplateParams.getTempParams(), t);
            String val = eval(tempStr, map).toString();
            if (isNumber && StringUtils.isNotEmpty(val)) {
                row.getCell(i + columnIndex).setCellValue(Double.parseDouble(val));
                row.getCell(i + columnIndex).setCellType(CellType.NUMERIC);
            } else {
                row.getCell(i + columnIndex).setCellValue(val);
            }
            row.getCell(i + columnIndex).setCellStyle(columns.get(i).getCellStyle());
            tempCreateCellSet.add(row.getRowNum() + "_" + (i + columnIndex));
        }

    }

    /**
     * Gets the value of the data for the iteration
     *
     * @param cell
     * @param name
     * @return
     */
    private List<ExcelTemplateParams> getAllDataColumns(Cell cell, String name) {
        List<ExcelTemplateParams> columns = new ArrayList<ExcelTemplateParams>();
        cell.setCellValue("");
        if (name.contains(END_STR)) {
            columns.add(new ExcelTemplateParams(name.replace(END_STR, EMPTY).trim(), cell.getCellStyle(), cell.getRow().getHeight()));
            return columns;
        }
        columns.add(new ExcelTemplateParams(name.trim(), cell.getCellStyle(), cell.getRow().getHeight()));
        int index = cell.getColumnIndex();
        //The number of columns
        int lastCellNum = cell.getRow().getLastCellNum();
        Cell tempCell;
        while (true) {
            tempCell = cell.getRow().getCell(++index);
            if (tempCell == null && index >= lastCellNum) {
                break;
            }
            String cellStringString;
            try {// Allow is empty, single means that it has been completed, because it may have been deleted
                cellStringString = tempCell.getStringCellValue();
                if (StringUtils.isBlank(cellStringString) && index >= lastCellNum) {
                    break;
                }
            } catch (Exception e) {
                throw new ExcelExportException("for each There is an empty string, please check the template");
            }
            // Leave the read cells empty
            tempCell.setCellValue("");
            if (cellStringString.contains(END_STR)) {
                columns.add(new ExcelTemplateParams(cellStringString.trim().replace(END_STR, ""), tempCell.getCellStyle(), tempCell.getRow().getHeight()));
                break;
            } else {
                if (cellStringString.trim().contains(teplateParams.getTempParams())) {
                    columns.add(new ExcelTemplateParams(cellStringString.trim(), tempCell.getCellStyle(), tempCell.getRow().getHeight()));
                } else if (cellStringString.trim().equals(EMPTY)) {
                    //A setting that may be a merged cell that allows empty data
                    columns.add(new ExcelTemplateParams(EMPTY, tempCell.getCellStyle(), tempCell.getRow().getHeight()));
                } else {
                    // The last line is deleted
                    break;
                }
            }

        }
        return columns;
    }

    /**
     * Sort and plug the exported sequence
     *
     * @param excelParams
     * @param titlemap
     * @return
     */
    private void sortAndFilterExportField(List<ExcelExportEntity> excelParams, Map<String, Integer> titlemap) {
        for (int i = excelParams.size() - 1; i >= 0; i--) {
            if (excelParams.get(i).getList() != null && excelParams.get(i).getList().size() > 0) {
                sortAndFilterExportField(excelParams.get(i).getList(), titlemap);
                if (excelParams.get(i).getList().size() == 0) {
                    excelParams.remove(i);
                } else {
                    excelParams.get(i).setOrderNum(i);
                }
            } else {
                if (titlemap.containsKey(excelParams.get(i).getName())) {
                    excelParams.get(i).setOrderNum(i);
                } else {
                    excelParams.remove(i);
                }
            }
        }
        sortAllParams(excelParams);
    }

    /**
     * Loop through the columns first, because there is a lot of data involved
     *
     * @param sheet
     * @param map
     */
    private void colForeach(Sheet sheet, Map<String, Object> map) throws Exception {
        Row row = null;
        Cell cell = null;
        int index = 0;
        while (index <= sheet.getLastRowNum()) {
            row = sheet.getRow(index++);
            if (row == null) {
                continue;
            }
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                cell = row.getCell(i);
                if (row.getCell(i) != null && (cell.getCellType() == CellType.STRING
                        || cell.getCellType() == CellType.NUMERIC)) {
                    String text = PoiCellUtil.getCellValue(cell);
                    if (text.contains(FOREACH_COL) || text.contains(FOREACH_COL_VALUE)) {
                        foreachCol(cell, map, text);
                    }
                }
            }
        }
    }

    /**
     * Loop through the list
     *
     * @param cell
     * @param map
     * @param name
     * @throws Exception
     */
    private void foreachCol(Cell cell, Map<String, Object> map, String name) throws Exception {
        boolean isCreate = name.contains(FOREACH_COL_VALUE);
        name = name.replace(FOREACH_COL_VALUE, EMPTY).replace(FOREACH_COL, EMPTY).replace(START_STR,
                EMPTY);
        String[] keys = name.replaceAll("\\s{1,}", " ").trim().split(" ");
        Collection<?> datas = (Collection<?>) PoiPublicUtil.getParamsValue(keys[0], map);
        Object[] columnsInfo = getAllDataColumns(cell, name.replace(keys[0], EMPTY),
                mergedRegionHelper);
        if (datas == null) {
            return;
        }
        Iterator<?> its = datas.iterator();
        int rowspan = (Integer) columnsInfo[0], colspan = (Integer) columnsInfo[1];
        @SuppressWarnings("unchecked")
        List<ExcelForEachParams> columns = (List<ExcelForEachParams>) columnsInfo[2];
        while (its.hasNext()) {
            Object t = its.next();
            setForeachRowCellValue(true, cell.getRow(), cell.getColumnIndex(), t, columns, map,
                    rowspan, colspan, mergedRegionHelper);
            if (cell.getRow().getCell(cell.getColumnIndex() + colspan) == null) {
                cell.getRow().createCell(cell.getColumnIndex() + colspan);
            }
            cell = cell.getRow().getCell(cell.getColumnIndex() + colspan);
        }
        if (isCreate) {
            cell = cell.getRow().getCell(cell.getColumnIndex() - 1);
            cell.setCellValue(cell.getStringCellValue() + END_STR);
        }
    }

    /**
     * Loop iteratively creates, traversing rows
     *
     * @param isCreate
     * @param row
     * @param columnIndex
     * @param t
     * @param columns
     * @param map
     * @param rowspan
     * @param colspan
     * @param mergedRegionHelper
     * @return rowSize, cellSize
     * @throws Exception
     */
    private int[] setForeachRowCellValue(boolean isCreate, Row row, int columnIndex, Object t,
                                         List<ExcelForEachParams> columns, Map<String, Object> map,
                                         int rowspan, int colspan,
                                         MergedRegionHelper mergedRegionHelper) throws Exception {
        createRowCellSetStyle(row, columnIndex, columns, rowspan, colspan);
        //Fill in the data
        ExcelForEachParams params;
        int loopSize = 1;
        int loopCi = 1;
        row = row.getSheet().getRow(row.getRowNum() - rowspan + 1);
        for (int k = 0; k < rowspan; k++) {
            int ci = columnIndex;
            row.setHeight(getMaxHeight(k, colspan, columns));
            for (int i = 0; i < colspan && i < columns.size(); i++) {
                boolean isNumber = false;
                params = columns.get(colspan * k + i);
                tempCreateCellSet.add(row.getRowNum() + "_" + (ci));
                if (params == null) {
                    continue;
                }
                if (StringUtils.isEmpty(params.getName())
                        && StringUtils.isEmpty(params.getConstValue())) {
                    row.getCell(ci).setCellStyle(params.getCellStyle());
                    ci = ci + params.getColspan();
                    continue;
                }
                String val;
                Object obj = null;
                //Is it a constant
                String tempStr = params.getName();
                if (StringUtils.isEmpty(params.getName())) {
                    val = params.getConstValue();
                } else {
                    if (isHasSymbol(tempStr, NUMBER_SYMBOL)) {
                        isNumber = true;
                        tempStr = tempStr.replaceFirst(NUMBER_SYMBOL, "");
                    }
                    map.put(teplateParams.getTempParams(), t);
                    boolean isDict = false;
                    String dict = null;
                    if (isHasSymbol(tempStr, DICT_HANDLER)) {
                        isDict = true;
                        dict = tempStr.substring(tempStr.indexOf(DICT_HANDLER) + 5).split(";")[0];
                        tempStr = tempStr.replaceFirst(DICT_HANDLER, "");
                        tempStr = tempStr.replaceFirst(dict + ";", "");
                    }
                    obj = eval(tempStr, map);
                    if (isDict && !(obj instanceof Collection)) {
                        obj = dictHandler.toName(dict, t, tempStr, obj);
                    }
                    val = obj.toString();
                }
                if (obj != null && obj instanceof Collection) {
                    // Need to find which level is the collection to facilitate the later replace
                    String collectName = evalFindName(tempStr, map);
                    int[] loop = setForEachLoopRowCellValue(row, ci, (Collection) obj, columns,
                            params, map, rowspan, colspan, mergedRegionHelper, collectName);
                    loopSize = Math.max(loopSize, loop[0]);
                    i += loop[1] - 1;
                    ci = loop[2] - params.getColspan();
                } else if (obj != null && obj instanceof ImageModel) {
                    ImageModel img = (ImageModel) obj;
                    row.getCell(ci).setCellValue("");
                    if (img.getRowspan() > 1 || img.getColspan() > 1) {
                        img.setHeight(0);
                        row.getCell(ci).getSheet().addMergedRegion(new CellRangeAddress(row.getCell(ci).getRowIndex(),
                                row.getCell(ci).getRowIndex() + img.getRowspan() - 1, row.getCell(ci).getColumnIndex(), row.getCell(ci).getColumnIndex() + img.getColspan() - 1));
                    }
                    createImageCell(row.getCell(ci), img.getHeight(), img.getRowspan(), img.getColspan(), img.getUrl(), img.getData());
                } else if (isNumber && StringUtils.isNotEmpty(val)) {
                    row.getCell(ci).setCellValue(Double.parseDouble(val));
                } else {
                    try {
                        row.getCell(ci).setCellValue(val);
                    } catch (Exception e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                }
                if (params.getCellStyle() != null) {
                    row.getCell(ci).setCellStyle(params.getCellStyle());
                }
                //If you merge cells, make the cell style the same as the previous one
                setMergedRegionStyle(row, ci, params);
                //Merge the corresponding cells
                if ((params.getRowspan() != 1 || params.getColspan() != 1)
                        && !mergedRegionHelper.isMergedRegion(row.getRowNum() + 1, ci)
                        && PoiCellUtil.isMergedRegion(row.getSheet(), row.getRowNum(), ci)) {
                    PoiMergeCellUtil.addMergedRegion(row.getSheet(), row.getRowNum(),
                            row.getRowNum() + params.getRowspan() - 1, ci,
                            ci + params.getColspan() - 1);
                }
                ci = ci + params.getColspan();
            }
            loopCi = Math.max(loopCi, ci);
            // The cells that need to be merged need to be merged--- the columns that are not the collection are merged
            if (loopSize > 1) {
                handlerLoopMergedRegion(row, columnIndex, columns, loopSize);
            }
            row = row.getSheet().getRow(row.getRowNum() + 1);
        }
        return new int[]{loopSize, loopCi};
    }

    /**
     * Image type Cell
     */
    public void createImageCell(Cell cell, double height, int rowspan, int colspan,
                                String imagePath, byte[] data) throws Exception {
        if (height > cell.getRow().getHeight()) {
            cell.getRow().setHeight((short) height);
        }
        ClientAnchor anchor;
        if (type.equals(ExcelType.HSSF)) {
            anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) cell.getColumnIndex(), cell.getRow().getRowNum(), (short) (cell.getColumnIndex() + colspan),
                    cell.getRow().getRowNum() + rowspan);
        } else {
            anchor = new XSSFClientAnchor(0, 0, 0, 0, (short) cell.getColumnIndex(), cell.getRow().getRowNum(), (short) (cell.getColumnIndex() + colspan),
                    cell.getRow().getRowNum() + rowspan);
        }
        if (StringUtils.isNotEmpty(imagePath)) {
            data = ImageCache.getImage(imagePath);
        }
        if (data != null) {
            PoiExcelGraphDataUtil.getDrawingPatriarch(cell.getSheet()).createPicture(anchor,
                    cell.getSheet().getWorkbook().addPicture(data, getImageType(data)));
        }
    }

    /**
     * Handle the inner loop
     *
     * @param row
     * @param columnIndex
     * @param obj
     * @param columns
     * @param params
     * @param map
     * @param rowspan
     * @param colspan
     * @param mergedRegionHelper
     * @param collectName
     * @return [rowNums, columnsNums, ciIndex]
     * @throws Exception
     */
    private int[] setForEachLoopRowCellValue(Row row, int columnIndex, Collection obj, List<ExcelForEachParams> columns,
                                             ExcelForEachParams params, Map<String, Object> map,
                                             int rowspan, int colspan,
                                             MergedRegionHelper mergedRegionHelper, String collectName) throws Exception {

        //Multiple traverses together - remove the first layer and go through all the data
        //STEP 1 GETS ALL THE SAME PROJECT FIELDS AS THE CURRENT ONE
        List<ExcelForEachParams> temp = getLoopEachParams(columns, columnIndex, collectName);
        Iterator<?> its = obj.iterator();
        Row tempRow = row;
        int nums = 0;
        int ci = columnIndex;
        while (its.hasNext()) {
            Object data = its.next();
            map.put("loop_" + columnIndex, data);
            int[] loopArr = setForeachRowCellValue(false, tempRow, columnIndex, data, temp, map, rowspan,
                    colspan, mergedRegionHelper);
            nums += loopArr[0];
            ci = Math.max(ci, loopArr[1]);
            map.remove("loop_" + columnIndex);
            tempRow = createRow(tempRow.getRowNum() + loopArr[0], row.getSheet(), false, rowspan);
        }
        for (int i = 0; i < temp.size(); i++) {
            temp.get(i).setName(temp.get(i).getTempName().pop());
            //It's all collections
            temp.get(i).setCollectCell(true);

        }
        return new int[]{nums, temp.size(), ci};
    }

    /**
     * Create and return the first Row
     *
     * @param sheet
     * @param rowIndex
     * @param isCreate
     * @param rows
     * @return
     */
    private Row createRow(int rowIndex, Sheet sheet, boolean isCreate, int rows) {
        for (int i = 0; i < rows; i++) {
            if (isCreate) {
                sheet.createRow(rowIndex++);
            } else if (sheet.getRow(rowIndex++) == null) {
                sheet.createRow(rowIndex - 1);
            }
        }
        return sheet.getRow(rowIndex - rows);
    }

    /**
     * According to the information that is currently a collection,
     * get the iteration of the entire collection and replace the prefix of the collection to facilitate the number later
     *
     * @param columns
     * @param columnIndex
     * @param collectName
     * @return
     */
    private List<ExcelForEachParams> getLoopEachParams(List<ExcelForEachParams> columns, int columnIndex, String collectName) {
        List<ExcelForEachParams> temp = new ArrayList<>();
        for (int i = 0; i < columns.size(); i++) {
            //Prepended to not be a collection
            columns.get(i).setCollectCell(false);
            if (columns.get(i) == null || columns.get(i).getName().contains(collectName)) {
                temp.add(columns.get(i));
                if (columns.get(i).getTempName() == null) {
                    columns.get(i).setTempName(new Stack<>());
                }
                columns.get(i).setCollectCell(true);
                columns.get(i).getTempName().push(columns.get(i).getName());
                columns.get(i).setName(columns.get(i).getName().replace(collectName, "loop_" + columnIndex));
            }
        }
        return temp;
    }

    /**
     * Style the rows
     *
     * @param row
     * @param columnIndex
     * @param columns
     * @param rowspan
     * @param colspan
     */
    private void createRowCellSetStyle(Row row, int columnIndex, List<ExcelForEachParams> columns,
                                       int rowspan, int colspan) {
        //All cells are created once
        for (int i = 0; i < rowspan; i++) {
            int size = columns.size();
            for (int j = columnIndex, max = columnIndex + colspan; j < max; j++) {
                if (row.getCell(j) == null) {
                    row.createCell(j);
                    CellStyle style = row.getRowNum() % 2 == 0
                            ? getStyles(false,
                            size <= j - columnIndex ? null : columns.get(j - columnIndex))
                            : getStyles(true,
                            size <= j - columnIndex ? null : columns.get(j - columnIndex));
                    //The returned styler is not empty when used, otherwise it is used with Excel settings, and the style set by Excel is more recommended
                    if (style != null) {
                        row.getCell(j).setCellStyle(style);
                    }
                }

            }
            if (i < rowspan - 1) {
                row = row.getSheet().getRow(row.getRowNum() + 1);
            }
        }
    }

    /**
     * Get Cell Style
     *
     * @param isSingle
     * @param excelForEachParams
     * @return
     */
    private CellStyle getStyles(boolean isSingle, ExcelForEachParams excelForEachParams) {
        return excelExportStyler.getTemplateStyles(isSingle, excelForEachParams);
    }

    /**
     * Gets the maximum height
     *
     * @param k
     * @param colspan
     * @param columns
     * @return
     */
    private short getMaxHeight(int k, int colspan, List<ExcelForEachParams> columns) {
        short high = columns.get(0).getHeight();
        int n = k;
        while (n > 0) {
            if (columns.get(n * colspan).getHeight() == 0) {
                n--;
            } else {
                high = columns.get(n * colspan).getHeight();
                break;
            }
        }
        return high;
    }

    private boolean isHasSymbol(String text, String symbol) {
        return text.startsWith(symbol) || text.contains("{" + symbol)
                || text.contains(" " + symbol);
    }

    /**
     * Iteration merges all data that is not a collection
     *
     * @param row
     * @param columnIndex
     * @param columns
     * @param loopSize
     */
    private void handlerLoopMergedRegion(Row row, int columnIndex, List<ExcelForEachParams> columns, int loopSize) {
        for (int i = 0; i < columns.size(); i++) {
            if (!columns.get(i).isCollectCell()) {
                PoiMergeCellUtil.addMergedRegion(row.getSheet(), row.getRowNum(),
                        row.getRowNum() + loopSize - 1, columnIndex,
                        columnIndex + columns.get(i).getColspan() - 1);
            }
            columnIndex = columnIndex + columns.get(i).getColspan();
        }
    }

    /**
     * Sets the style of the merged cells
     *
     * @param row
     * @param ci
     * @param params
     */
    private void setMergedRegionStyle(Row row, int ci, ExcelForEachParams params) {
        //The first row of data
        for (int i = 1; i < params.getColspan(); i++) {
            if (params.getCellStyle() != null) {
                row.getCell(ci + i).setCellStyle(params.getCellStyle());
            }
        }
        for (int i = 1; i < params.getRowspan(); i++) {
            for (int j = 0; j < params.getColspan(); j++) {
                if (params.getCellStyle() != null) {
                    row.getCell(ci + j).setCellStyle(params.getCellStyle());
                }
            }
        }
    }

    /**
     * Gets the value of the data for the iteration
     *
     * @param cell
     * @param name
     * @param mergedRegionHelper
     * @return
     */
    private Object[] getAllDataColumns(Cell cell, String name,
                                       MergedRegionHelper mergedRegionHelper) {
        List<ExcelForEachParams> columns = new ArrayList<ExcelForEachParams>();
        cell.setCellValue("");
        columns.add(getExcelTemplateParams(name.replace(END_STR, EMPTY), cell, mergedRegionHelper));
        int rowspan = 1, colspan = 1;
        if (!name.contains(END_STR)) {
            int index = cell.getColumnIndex();
            //Save the start column of col
            int startIndex = cell.getColumnIndex();
            Row row = cell.getRow();
            while (index < row.getLastCellNum()) {
                int colSpan = columns.get(columns.size() - 1) != null
                        ? columns.get(columns.size() - 1).getColspan() : 1;
                index += colSpan;


                for (int i = 1; i < colSpan; i++) {
                    //Add merged cells, which may not be empty but have no value, so they also need to be skipped
                    columns.add(null);
                    continue;
                }
                cell = row.getCell(index);
                //Probably merged cells
                if (cell == null) {
                    //Read is judgment, skip
                    columns.add(null);
                    continue;
                }
                String cellStringString;
                try {//Do not allow null Convenience cells must have an end and a value
                    cellStringString = cell.getStringCellValue();
                    if (StringUtils.isBlank(cellStringString) && colspan + startIndex <= index) {
                        throw new ExcelExportException("There is an empty string in for each, please check the template");
                    } else if (StringUtils.isBlank(cellStringString)
                            && colspan + startIndex > index) {
                        //Read is judgment, skip, data is empty, but not the first time reading this column, so it can be skipped
                        columns.add(new ExcelForEachParams(null, cell.getCellStyle(), (short) 0));
                        continue;
                    }
                } catch (Exception e) {
                    throw new ExcelExportException(ExcelExportEnum.TEMPLATE_ERROR, e);
                }
                //Leave the read cells empty
                cell.setCellValue("");
                if (cellStringString.contains(END_STR)) {
                    columns.add(getExcelTemplateParams(cellStringString.replace(END_STR, EMPTY),
                            cell, mergedRegionHelper));
                    //Complete missing cells (after merged cells)
                    int lastCellColspan = columns.get(columns.size() - 1).getColspan();
                    for (int i = 1; i < lastCellColspan; i++) {
                        //Add merged cells, which may not be empty but have no value, so they also need to be skipped
                        columns.add(null);
                    }
                    break;
                } else if (cellStringString.contains(WRAP)) {
                    columns.add(getExcelTemplateParams(cellStringString.replace(WRAP, EMPTY), cell,
                            mergedRegionHelper));
                    //Discover line breaks and perform line breaks
                    colspan = index - startIndex + 1;
                    index = startIndex - columns.get(columns.size() - 1).getColspan();
                    row = row.getSheet().getRow(row.getRowNum() + 1);
                    rowspan++;
                } else {
                    columns.add(getExcelTemplateParams(cellStringString.replace(WRAP, EMPTY), cell,
                            mergedRegionHelper));
                }
            }
        }
        colspan = 0;
        for (int i = 0; i < columns.size(); i++) {
            colspan += columns.get(i) != null ? columns.get(i).getColspan() : 0;
        }
        colspan = colspan / rowspan;
        return new Object[]{rowspan, colspan, columns};
    }

    /**
     * Gets the template parameters
     *
     * @param name
     * @param cell
     * @param mergedRegionHelper
     * @return
     */
    private ExcelForEachParams getExcelTemplateParams(String name, Cell cell,
                                                      MergedRegionHelper mergedRegionHelper) {
        name = name.trim();
        ExcelForEachParams params = new ExcelForEachParams(name, cell.getCellStyle(),
                cell.getRow().getHeight());
        //Determine if it is a constant
        if (name.startsWith(CONST) && name.endsWith(CONST)) {
            params.setName(null);
            params.setConstValue(name.substring(1, name.length() - 1));
        }
        //Determine if it is empty
        if (NULL.equals(name)) {
            params.setName(null);
            params.setConstValue(EMPTY);
        }
        //Gets the data of the merged cell
        if (mergedRegionHelper.isMergedRegion(cell.getRowIndex() + 1, cell.getColumnIndex())) {
            Integer[] colAndrow = mergedRegionHelper.getRowAndColSpan(cell.getRowIndex() + 1,
                    cell.getColumnIndex());
            params.setRowspan(colAndrow[0]);
            params.setColspan(colAndrow[1]);
        }
        return params;
    }
}
