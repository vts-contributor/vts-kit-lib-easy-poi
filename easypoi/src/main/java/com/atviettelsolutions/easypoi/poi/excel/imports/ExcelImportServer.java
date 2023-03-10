
package com.atviettelsolutions.easypoi.poi.excel.imports;

import com.atviettelsolutions.easypoi.core.util.ApplicationContextUtil;
import com.atviettelsolutions.easypoi.poi.common.ExcelUtil;
import com.atviettelsolutions.easypoi.poi.excel.imports.base.ImportFileServiceI;
import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import com.atviettelsolutions.easypoi.poi.exeption.ExcelImportException;
import com.atviettelsolutions.easypoi.poi.exeption.excel.enums.ExcelImportEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelTarget;
import com.atviettelsolutions.easypoi.poi.excel.entity.ImportParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelCollectionParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelImportEntity;
import com.atviettelsolutions.easypoi.poi.excel.entity.result.ExcelImportResult;
import com.atviettelsolutions.easypoi.poi.excel.entity.result.ExcelVerifyHanlderResult;
import com.atviettelsolutions.easypoi.poi.excel.imports.base.ImportBaseService;
import com.atviettelsolutions.easypoi.poi.excel.imports.verifys.VerifyHandlerServer;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Excel Import the service
 *
 * @author caprocute
 */
@SuppressWarnings({"rawtypes", "unchecked", "hiding"})
public class ExcelImportServer extends ImportBaseService {

    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelImportServer.class);

    private CellValueServer cellValueServer;

    private VerifyHandlerServer verifyHandlerServer;

    private boolean verfiyFail = false;
    /**
     * Exception data styler
     */
    private CellStyle errorCellStyle;

    public ExcelImportServer() {
        this.cellValueServer = new CellValueServer();
        this.verifyHandlerServer = new VerifyHandlerServer();
    }

    /***
     * Continue adding elements to the List
     *
     * @param object
     * @param param
     * @param row
     * @param titlemap
     * @param targetId
     * @param pictures
     * @param params
     */
    private void addListContinue(Object object, ExcelCollectionParams param, Row row, Map<Integer, String> titlemap, String targetId, Map<String, PictureData> pictures, ImportParams params) throws Exception {
        Collection collection = (Collection) PoiPublicUtil.getMethod(param.getName(), object.getClass()).invoke(object, new Object[]{});
        Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
        String picId;
        boolean isUsed = false;// Whether this object needs to be added
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            String titleString = (String) titlemap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                if (param.getExcelParams().get(titleString).getType() == 2) {
                    picId = row.getRowNum() + "_" + i;
                    saveImage(object, picId, param.getExcelParams(), titleString, pictures, params);
                } else {
                    saveFieldValue(params, entity, cell, param.getExcelParams(), titleString, row);
                }
                isUsed = true;
            }
        }
        if (isUsed) {
            collection.add(entity);
        }
    }

    /**
     * Get the value of the key, and get different values for different types
     *
     * @param cell
     * @return
     * @author caprocute
     * @date 2013-11-21
     */
    private String getKeyValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        Object obj = null;
        switch (cell.getCellTypeEnum()) {
            case STRING:
                obj = cell.getStringCellValue();
                break;
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case NUMERIC:
                obj = cell.getNumericCellValue();
                break;
            case FORMULA:
                obj = cell.getCellFormula();
                break;
        }
        return obj == null ? null : obj.toString().trim();
    }

    /**
     * Get the saved real path
     *
     * @param excelImportEntity
     * @param object
     * @return
     * @throws Exception
     */
    private String getSaveUrl(ExcelImportEntity excelImportEntity, Object object) throws Exception {
        String url = "";
        if (excelImportEntity.getSaveUrl().equals("upload")) {
            if (excelImportEntity.getMethods() != null && excelImportEntity.getMethods().size() > 0) {
                object = getFieldBySomeMethod(excelImportEntity.getMethods(), object);
            }
            url = object.getClass().getName().split("\\.")[object.getClass().getName().split("\\.").length - 1];
            return excelImportEntity.getSaveUrl() + "/" + url.substring(0, url.lastIndexOf("Entity"));
        }
        return excelImportEntity.getSaveUrl();
    }

    private <T> List<T> importExcel(Collection<T> result, Sheet sheet, Class<?> pojoClass, ImportParams params, Map<String, PictureData> pictures) throws Exception {
        List collection = new ArrayList();
        Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();
        List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();
        String targetId = null;
        if (!Map.class.equals(pojoClass)) {
            Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            if (etarget != null) {
                targetId = etarget.value();
            }
            getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
        }
        ignoreHeaderHandler(excelParams, params);
        Iterator<Row> rows = sheet.rowIterator();
        Map<Integer, String> titlemap = getTitleMap(sheet, rows, params, excelCollection);
        Set<String> keys = excelParams.keySet();
        for (String key : keys) {
            if (key.startsWith("FIXED_")) {
                String[] arr = key.split("_");
                titlemap.put(Integer.parseInt(arr[1]), key);
            }
        }
        Set<Integer> columnIndexSet = titlemap.keySet();
        Integer maxColumnIndex = Collections.max(columnIndexSet);
        Integer minColumnIndex = Collections.min(columnIndexSet);
        Row row = null;
        //Skip header and header rows
        for (int j = 0; j < params.getTitleRows() + params.getHeadRows(); j++) {
            row = rows.next();
        }
        Object object = null;
        String picId;
        while (rows.hasNext() && (row == null || sheet.getLastRowNum() - row.getRowNum() > params.getLastOfInvalidRow())) {
            row = rows.next();
            // Determine whether it is a collection element or not, if so, continue to join the collection, or create a new object
            Cell keyIndexCell = row.getCell(params.getKeyIndex());
            if (excelCollection.size() > 0 && StringUtils.isEmpty(getKeyValue(keyIndexCell)) && object != null && !Map.class.equals(pojoClass)) {
                for (ExcelCollectionParams param : excelCollection) {
                    addListContinue(object, param, row, titlemap, targetId, pictures, params);
                }

            } else {
                object = PoiPublicUtil.createObject(pojoClass, targetId);
                try {
                    int firstCellNum = row.getFirstCellNum();
                    if (firstCellNum > minColumnIndex) {
                        firstCellNum = minColumnIndex;
                    }
                    int lastCellNum = row.getLastCellNum();
                    if (lastCellNum < maxColumnIndex + 1) {
                        lastCellNum = maxColumnIndex + 1;
                    }
                    for (int i = firstCellNum, le = lastCellNum; i < le; i++) {
                        Cell cell = row.getCell(i);
                        String titleString = (String) titlemap.get(i);
                        if (excelParams.containsKey(titleString) || Map.class.equals(pojoClass)) {
                            if (excelParams.get(titleString) != null && excelParams.get(titleString).getType() == 2) {
                                picId = row.getRowNum() + "_" + i;
                                saveImage(object, picId, excelParams, titleString, pictures, params);
                            } else {
                                if (params.getImageList() != null && params.getImageList().contains(titleString)) {
                                    if (pictures != null) {
                                        picId = row.getRowNum() + "_" + i;
                                        PictureData image = pictures.get(picId);
                                        if (image != null) {
                                            byte[] data = image.getData();
                                            params.getDataHanlder().setMapValue((Map) object, titleString, data);
                                        }
                                    }
                                } else {
                                    saveFieldValue(params, object, cell, excelParams, titleString, row);
                                }
                            }
                        }
                    }

                    for (ExcelCollectionParams param : excelCollection) {
                        addListContinue(object, param, row, titlemap, targetId, pictures, params);
                    }
                    if (isNotNullObject(pojoClass, object)) {
                        collection.add(object);
                    }
                } catch (ExcelImportException e) {
                    if (!e.getType().equals(ExcelImportEnum.VERIFY_ERROR)) {
                        throw new ExcelImportException(e.getType(), e);
                    }
                }
            }
        }
        return collection;
    }

    /**
     * Determines that the current object is not empty
     *
     * @param pojoClass
     * @param object
     * @return
     */
    private boolean isNotNullObject(Class pojoClass, Object object) {
        try {
            Method method = pojoClass.getMethod("isNullObject");
            if (method != null) {
                Object flag = method.invoke(object);
                if (flag != null && true == Boolean.parseBoolean(flag.toString())) {
                    return false;
                }
            }
        } catch (NoSuchMethodException e) {
            LOGGER.debug("No method defined isNullObject");
        } catch (IllegalAccessException e) {
            LOGGER.warn("You do not have permission to access the method isNullObject");
        } catch (InvocationTargetException e) {
            LOGGER.warn("The method call failed isNullObject");
        }
        return true;
    }

    /**
     * Gets the ignored header information
     *
     * @param excelParams
     * @param params
     */
    private void ignoreHeaderHandler(Map<String, ExcelImportEntity> excelParams, ImportParams params) {
        List<String> ignoreList = new ArrayList<>();
        for (String key : excelParams.keySet()) {
            String temp = excelParams.get(key).getGroupName();
            if (temp != null && temp.length() > 0) {
                ignoreList.add(temp);
            }
        }
        params.setIgnoreHeaderList(ignoreList);
    }

    /**
     * Gets the corresponding information for the column names of table fields
     *
     * @param rows
     * @param params
     * @param excelCollection
     * @return
     */
    private Map<Integer, String> getTitleMap(Sheet sheet, Iterator<Row> rows, ImportParams params, List<ExcelCollectionParams> excelCollection) throws Exception {
        Map<Integer, String> titlemap = new HashMap<Integer, String>();
        Iterator<Cell> cellTitle = null;
        String collectionName = null;
        ExcelCollectionParams collectionParams = null;
        Row headRow = null;
        int headBegin = params.getTitleRows();
        int allRowNum = sheet.getPhysicalNumberOfRows();
        //To find the first row of headers, each sheet must have at least one row of headers
        while (headRow == null && headBegin < allRowNum) {
            headRow = sheet.getRow(headBegin++);
        }
        if (headRow == null) {
            throw new Exception("The file is not recognized");
        }

        //Set the number of header rows
        if (ExcelUtil.isMergedRegion(sheet, headRow.getRowNum(), 0)) {
            params.setHeadRows(2);
        } else {
            params.setHeadRows(1);
        }
        cellTitle = headRow.cellIterator();
        while (cellTitle.hasNext()) {
            Cell cell = cellTitle.next();
            String value = getKeyValue(cell);
            if (StringUtils.isNotEmpty(value)) {
                titlemap.put(cell.getColumnIndex(), value);//Add to the header list
            }
        }

        //Multi-row headers
        for (int j = headBegin; j < headBegin + params.getHeadRows() - 1; j++) {
            headRow = sheet.getRow(j);
            cellTitle = headRow.cellIterator();
            while (cellTitle.hasNext()) {
                Cell cell = cellTitle.next();
                String value = getKeyValue(cell);
                if (StringUtils.isNotEmpty(value)) {
                    int columnIndex = cell.getColumnIndex();
                    //Whether the previous row of the current cell is a merged cell
                    if (ExcelUtil.isMergedRegion(sheet, cell.getRowIndex() - 1, columnIndex)) {
                        collectionName = ExcelUtil.getMergedRegionValue(sheet, cell.getRowIndex() - 1, columnIndex);
                        if (params.isIgnoreHeader(collectionName)) {
                            titlemap.put(cell.getColumnIndex(), value);
                        } else {
                            titlemap.put(cell.getColumnIndex(), collectionName + "_" + value);
                        }
                    } else {
                        // There is another special scenario when the previous row is not merged:
                        // if the current cell and the cell above are in the same column, that is,
                        // there is only one subtable field, so the header does not appear across columns
                        String prefixTitle = titlemap.get(cell.getColumnIndex());
                        if (prefixTitle != null && !"".equals(prefixTitle)) {
                            titlemap.put(cell.getColumnIndex(), prefixTitle + "_" + value);
                        } else {
                            titlemap.put(cell.getColumnIndex(), value);
                        }
                    }

                }
            }
        }
        return titlemap;
    }

    /**
     * Gets the collection information corresponding to this name
     *
     * @param excelCollection
     * @param collectionName
     * @return
     */
    private ExcelCollectionParams getCollectionParams(List<ExcelCollectionParams> excelCollection, String collectionName) {
        for (ExcelCollectionParams excelCollectionParams : excelCollection) {
            if (collectionName.equals(excelCollectionParams.getExcelName())) {
                return excelCollectionParams;
            }
        }
        return null;
    }

    /**
     * Excel imports field field types Integer, Long, Double, Date, String, Boolean
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public ExcelImportResult importExcelByIs(InputStream inputstream, Class<?> pojoClass, ImportParams params) throws Exception {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel import start ,class is {}", pojoClass);
        }
        List<T> result = new ArrayList<T>();
        Workbook book = null;
        boolean isXSSFWorkbook = false;
        if (!(inputstream.markSupported())) {
            inputstream = new PushbackInputStream(inputstream, 8);
        }

        book = WorkbookFactory.create(inputstream);
        if (book instanceof XSSFWorkbook) {
            isXSSFWorkbook = true;
        }
        LOGGER.info("  >>>  poi3 upgrade to 4.0 compatible retrofit work, isXSSFWorkbook = " + isXSSFWorkbook);
          if (params.getSheetNum() == 0) {
            int sheetNum = book.getNumberOfSheets();
            if (sheetNum > 0) {
                params.setSheetNum(sheetNum);
            }
        }
      createErrorCellStyle(book);
        Map<String, PictureData> pictures;

        for (int i = params.getStartSheetIndex(); i < params.getStartSheetIndex()
                + params.getSheetNum(); i++) {
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" start to read excel by is ,startTime is {}", System.currentTimeMillis());
            }
            if (isXSSFWorkbook) {
                pictures = PoiPublicUtil.getSheetPictrues07((XSSFSheet) book.getSheetAt(i), (XSSFWorkbook) book);
            } else {
                pictures = PoiPublicUtil.getSheetPictrues03((HSSFSheet) book.getSheetAt(i), (HSSFWorkbook) book);
            }
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" end to read excel by is ,endTime is {}", new Date().getTime());
            }
            result.addAll(importExcel(result, book.getSheetAt(i), pojoClass, params, pictures));
            if (LOGGER.isDebugEnabled()) {
                LOGGER.debug(" end to read excel list by pos ,endTime is {}", new Date().getTime());
            }
        }
        if (params.isNeedSave()) {
            saveThisExcel(params, pojoClass, isXSSFWorkbook, book);
        }
        return new ExcelImportResult(result, verfiyFail, book);
    }

    /**
     * @param is
     * @return
     * @throws IOException
     */
    public static byte[] getBytes(InputStream is) throws IOException {
        ByteArrayOutputStream buffer = new ByteArrayOutputStream();

        int len;
        byte[] data = new byte[100000];
        while ((len = is.read(data, 0, data.length)) != -1) {
            buffer.write(data, 0, len);
        }

        buffer.flush();
        return buffer.toByteArray();
    }

    /**
     * Save field values (get values, check values, append error messages)
     *
     * @param params
     * @param object
     * @param cell
     * @param excelParams
     * @param titleString
     * @param row
     * @throws Exception
     */
    private void saveFieldValue(ImportParams params, Object object, Cell cell, Map<String, ExcelImportEntity> excelParams, String titleString, Row row) throws Exception {
        Object value = cellValueServer.getValue(params.getDataHanlder(), object, cell, excelParams, titleString);
        if (object instanceof Map) {
            if (params.getDataHanlder() != null) {
                params.getDataHanlder().setMapValue((Map) object, titleString, value);
            } else {
                ((Map) object).put(titleString, value);
            }
        } else {
            ExcelVerifyHanlderResult verifyResult = verifyHandlerServer.verifyData(object, value, titleString, excelParams.get(titleString).getVerify(), params.getVerifyHanlder());
            if (verifyResult.isSuccess()) {
                setValues(excelParams.get(titleString), object, value);
            } else {
                Cell errorCell = row.createCell(row.getLastCellNum());
                errorCell.setCellValue(verifyResult.getMsg());
                errorCell.setCellStyle(errorCellStyle);
                verfiyFail = true;
                throw new ExcelImportException(ExcelImportEnum.VERIFY_ERROR);
            }
        }
    }

    /**
     * @param object
     * @param picId
     * @param excelParams
     * @param titleString
     * @param pictures
     * @param params
     * @throws Exception
     */
    private void saveImage(Object object, String picId, Map<String, ExcelImportEntity> excelParams, String titleString, Map<String, PictureData> pictures, ImportParams params) throws Exception {
        if (pictures == null || pictures.get(picId) == null) {
            return;
        }
        PictureData image = pictures.get(picId);
        byte[] data = image.getData();
        String fileName = "pic" + Math.round(Math.random() * 100000000000L);
        fileName += "." + PoiPublicUtil.getFileExtendName(data);
        int saveType = excelParams.get(titleString).getSaveType();
        if (saveType == 1) {
            String path = PoiPublicUtil.getWebRootPath(getSaveUrl(excelParams.get(titleString), object));
            File savefile = new File(path);
            if (!savefile.exists()) {
                savefile.mkdirs();
            }
            savefile = new File(path + "/" + fileName);
            FileOutputStream fos = new FileOutputStream(savefile);
            fos.write(data);
            fos.close();
            setValues(excelParams.get(titleString), object, getSaveUrl(excelParams.get(titleString), object) + "/" + fileName);
        } else if (saveType == 2) {
            setValues(excelParams.get(titleString), object, data);
        } else {
            ImportFileServiceI importFileService = null;
            try {
                importFileService = ApplicationContextUtil.getContext().getBean(ImportFileServiceI.class);
            } catch (Exception e) {
                System.err.println(e.getMessage());
            }
            if (importFileService != null) {
                String dbPath = importFileService.doUpload(data);
                setValues(excelParams.get(titleString), object, dbPath);
            }
        }
    }

    private void createErrorCellStyle(Workbook workbook) {
        errorCellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(Font.COLOR_RED);
        errorCellStyle.setFont(font);
    }

}
