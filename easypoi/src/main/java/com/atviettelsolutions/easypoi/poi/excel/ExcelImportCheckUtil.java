package com.atviettelsolutions.easypoi.poi.excel;

import com.atviettelsolutions.easypoi.core.util.ApplicationContextUtil;
import com.atviettelsolutions.easypoi.dict.service.AutoPoiDictServiceI;
import com.atviettelsolutions.easypoi.poi.common.ExcelUtil;
import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import com.atviettelsolutions.easypoi.poi.exeption.ExcelImportException;
import com.atviettelsolutions.easypoi.poi.exeption.excel.enums.ExcelImportEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import com.atviettelsolutions.easypoi.poi.excel.annotation.Excel;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelCollection;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelTarget;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelVerify;
import com.atviettelsolutions.easypoi.poi.excel.entity.ImportParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelCollectionParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelImportEntity;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelVerifyEntity;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.io.PushbackInputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.math.BigDecimal;
import java.util.*;

/**
 * EXCEL INCLUE CHECK
 * Verify that the Excel title exists, and the current default is 0.8 (80%) to pass the verification
 */
public class ExcelImportCheckUtil {
    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelImportCheckUtil.class);

    /**
     * When there are titles that arrive how much can pass verification
     */
    public static final Double defScreenRate = 0.8;

    /**
     * check inclue filed match rate
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @return
     */
    public static Boolean check(InputStream inputstream, Class<?> pojoClass, ImportParams params) {
        return check(inputstream, pojoClass, params, defScreenRate);
    }

    /**
     * check inclue filed match rate
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @param screenRate  field match rate (defalut:0.8)
     * @return
     */
    public static Boolean check(InputStream inputstream, Class<?> pojoClass, ImportParams params, Double screenRate) {
        Workbook book = null;
        int errorNum = 0;
        int successNum = 0;
        if (!(inputstream.markSupported())) {
            inputstream = new PushbackInputStream(inputstream, 8);
        }
        try {
            book = WorkbookFactory.create(inputstream);
            LOGGER.info("  >>>  POI 3 upgraded to 4 compatible retrofit work, pojoClass=" + pojoClass);
        } catch (Exception e) {
            e.printStackTrace();
        }
        for (int i = 0; i < params.getSheetNum(); i++) {
            Row row = null;
            //Skip header and header rows
            Iterator<Row> rows;
            try {
                rows = book.getSheetAt(i).rowIterator();
            } catch (Exception e) {
                //Empty descriptions cannot be read, so they are not excel
                throw new RuntimeException("Please import the Excel file in the correct format！");
            }


            for (int j = 0; j < params.getTitleRows() + params.getHeadRows(); j++) {
                try {
                    row = rows.next();
                } catch (NoSuchElementException e) {
                    //Empty caption is not out and Excel format is incorrect
                    throw new RuntimeException("Please fill in the content title！");
                }
            }
            Sheet sheet = book.getSheetAt(i);
            Map<Integer, String> titlemap = null;
            try {
                titlemap = getTitleMap(sheet, params);
            } catch (Exception e) {
                e.printStackTrace();
            }
            Set<Integer> columnIndexSet = titlemap.keySet();
            Integer maxColumnIndex = Collections.max(columnIndexSet);
            Integer minColumnIndex = Collections.min(columnIndexSet);
            while (rows.hasNext() && (row == null || sheet.getLastRowNum() - row.getRowNum() > params.getLastOfInvalidRow())) {
                row = rows.next();
                Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();
                List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();
                String targetId = null;
                if (!Map.class.equals(pojoClass)) {
                    Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
                    ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
                    if (etarget != null) {
                        targetId = etarget.value();
                    }
                    try {
                        getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                try {
                    int firstCellNum = row.getFirstCellNum();
                    if (firstCellNum > minColumnIndex) {
                        firstCellNum = minColumnIndex;
                    }
                    int lastCellNum = row.getLastCellNum();
                    if (lastCellNum < maxColumnIndex + 1) {
                        lastCellNum = maxColumnIndex + 1;
                    }
                    for (int j = firstCellNum, le = lastCellNum; j < le; j++) {
                        String titleString = (String) titlemap.get(j);
                        if (excelParams.containsKey(titleString) || Map.class.equals(pojoClass)) {
                            successNum += 1;
                        } else {
                            if (excelCollection.size() > 0) {
                                Iterator var33 = excelCollection.iterator();
                                ExcelCollectionParams param = (ExcelCollectionParams) var33.next();
                                if (param.getExcelParams().containsKey(titleString)) {
                                    successNum += 1;
                                } else {
                                    errorNum += 1;
                                }
                            } else {
                                errorNum += 1;
                            }
                        }
                    }
                    if (successNum < errorNum) {
                        return false;
                    } else if (successNum > errorNum) {
                        if (errorNum > 0) {
                            double newNumber = (double) successNum / (successNum + errorNum);
                            BigDecimal bg = new BigDecimal(newNumber);
                            double f1 = bg.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue();
                            if (f1 < screenRate) {
                                return false;
                            } else {
                                return true;
                            }
                        } else {
                            return true;
                        }
                    } else if (successNum == errorNum) {
                        return false;
                    } else {
                        return false;
                    }
                } catch (ExcelImportException e) {
                    if (!e.getType().equals(ExcelImportEnum.VERIFY_ERROR)) {
                        throw new ExcelImportException(e.getType(), e);
                    }
                }

            }
        }
        return null;
    }

    /**
     * Gets the file name title
     *
     * @throws Exception
     * @author caprocute
     */
    private static Map<Integer, String> getTitleMap(Sheet sheet, ImportParams params) throws Exception {
        Map<Integer, String> titlemap = new HashMap<Integer, String>();
        Iterator<Cell> cellTitle = null;
        String collectionName = null;
        Row headRow = null;
        int headBegin = params.getTitleRows();
        int allRowNum = sheet.getPhysicalNumberOfRows();
        while (headRow == null && headBegin < allRowNum) {
            headRow = sheet.getRow(headBegin++);
        }
        if (headRow == null) {
            throw new Exception("The file is not recognized");
        }
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
                        titlemap.put(cell.getColumnIndex(), value);
                    }
                }
            }
        }
        return titlemap;
    }

    /**
     * Get the value of the key, and get different values for different types
     *
     * @param cell
     * @return
     * @author caprocute
     */
    private static String getKeyValue(Cell cell) {
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
     * Get all the fields that need to be exported
     *
     * @param targetId
     * @param fields
     * @param excelCollection
     * @throws Exception
     */
    public static void getAllExcelField(String targetId, Field[] fields, Map<String, ExcelImportEntity> excelParams, List<ExcelCollectionParams> excelCollection, Class<?> pojoClass, List<Method> getMethods) throws Exception {
        ExcelImportEntity excelEntity = null;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (PoiPublicUtil.isNotUserExcelUserThis(null, field, targetId)) {
                continue;
            }
            if (PoiPublicUtil.isCollection(field.getType())) {
                // Collection object setting properties
                ExcelCollectionParams collection = new ExcelCollectionParams();
                collection.setName(field.getName());
                Map<String, ExcelImportEntity> temp = new HashMap();
                ParameterizedType pt = (ParameterizedType) field.getGenericType();
                Class<?> clz = (Class) pt.getActualTypeArguments()[0];
                collection.setType(clz);
                getExcelFieldList(targetId, PoiPublicUtil.getClassFields(clz), clz, temp, (List) null);
                collection.setExcelParams(temp);
                collection.setExcelName(((ExcelCollection) field.getAnnotation(ExcelCollection.class)).name());
                additionalCollectionName(collection);
                excelCollection.add(collection);
            } else if (PoiPublicUtil.isJavaClass(field)) {
                addEntityToMap(targetId, field, (ExcelImportEntity) excelEntity, pojoClass, getMethods, excelParams);
            } else {
                List<Method> newMethods = new ArrayList<Method>();
                if (getMethods != null) {
                    newMethods.addAll(getMethods);
                }
                newMethods.add(PoiPublicUtil.getMethod(field.getName(), pojoClass));
                getAllExcelField(targetId, PoiPublicUtil.getClassFields(field.getType()), excelParams, excelCollection, field.getType(), newMethods);
            }
        }
    }

    public static void getExcelFieldList(String targetId, Field[] fields, Class<?> pojoClass, Map<String, ExcelImportEntity> temp, List<Method> getMethods) throws Exception {
        ExcelImportEntity excelEntity = null;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (!PoiPublicUtil.isNotUserExcelUserThis((List) null, field, targetId)) {
                if (PoiPublicUtil.isJavaClass(field)) {
                    addEntityToMap(targetId, field, (ExcelImportEntity) excelEntity, pojoClass, getMethods, temp);
                } else {
                    List<Method> newMethods = new ArrayList();
                    if (getMethods != null) {
                        newMethods.addAll(getMethods);
                    }

                    newMethods.add(PoiPublicUtil.getMethod(field.getName(), pojoClass, field.getType()));
                    getExcelFieldList(targetId, PoiPublicUtil.getClassFields(field.getType()), field.getType(), temp, newMethods);
                }
            }
        }
    }

    /**
     * Append the collection name to the front
     *
     * @param collection
     */
    private static void additionalCollectionName(ExcelCollectionParams collection) {
        Set<String> keys = new HashSet();
        keys.addAll(collection.getExcelParams().keySet());
        Iterator var3 = keys.iterator();

        while (var3.hasNext()) {
            String key = (String) var3.next();
            collection.getExcelParams().put(collection.getExcelName() + "_" + key, collection.getExcelParams().get(key));
            collection.getExcelParams().remove(key);
        }
    }

    /**
     * Put this annotation parsing in the type object
     *
     * @param targetId
     * @param field
     * @param excelEntity
     * @param pojoClass
     * @param getMethods
     * @param temp
     * @throws Exception
     */
    public static void addEntityToMap(String targetId, Field field, ExcelImportEntity excelEntity, Class<?> pojoClass, List<Method> getMethods, Map<String, ExcelImportEntity> temp) throws Exception {
        Excel excel = field.getAnnotation(Excel.class);
        excelEntity = new ExcelImportEntity();
        excelEntity.setType(excel.type());
        excelEntity.setSaveUrl(excel.savePath());
        excelEntity.setSaveType(excel.imageType());
        excelEntity.setReplace(excel.replace());
        excelEntity.setDatabaseFormat(excel.databaseFormat());
        excelEntity.setVerify(getImportVerify(field));
        excelEntity.setSuffix(excel.suffix());
        excelEntity.setNumFormat(excel.numFormat());
        excelEntity.setGroupName(excel.groupName());
        excelEntity.setMultiReplace(excel.multiReplace());
        if (StringUtils.isNotEmpty(excel.dicCode())) {
            AutoPoiDictServiceI jeecgDictService = null;
            try {
                jeecgDictService = ApplicationContextUtil.getContext().getBean(AutoPoiDictServiceI.class);
            } catch (Exception e) {
            }
            if (jeecgDictService != null) {
                String[] dictReplace = jeecgDictService.queryDict(excel.dictTable(), excel.dicCode(), excel.dicText());
                if (excelEntity.getReplace() != null && dictReplace != null && dictReplace.length != 0) {
                    excelEntity.setReplace(dictReplace);
                }
            }
        }
        getExcelField(targetId, field, excelEntity, excel, pojoClass);
        if (getMethods != null) {
            List<Method> newMethods = new ArrayList<Method>();
            newMethods.addAll(getMethods);
            newMethods.add(excelEntity.getMethod());
            excelEntity.setMethods(newMethods);
        }
        temp.put(excelEntity.getName(), excelEntity);

    }

    public static void getExcelField(String targetId, Field field, ExcelImportEntity excelEntity, Excel excel, Class<?> pojoClass) throws Exception {
        excelEntity.setName(getExcelName(excel.name(), targetId));
        String fieldname = field.getName();
        excelEntity.setMethod(PoiPublicUtil.getMethod(fieldname, pojoClass, field.getType(), excel.importConvert()));
        if (StringUtils.isNotEmpty(excel.importFormat())) {
            excelEntity.setFormat(excel.importFormat());
        } else {
            excelEntity.setFormat(excel.format());
        }
    }

    /**
     * Determines the name displayed in this cell
     *
     * @param exportName
     * @param targetId
     * @return
     */
    public static String getExcelName(String exportName, String targetId) {
        if (exportName.indexOf("_") < 0) {
            return exportName;
        }
        String[] arr = exportName.split(",");
        for (String str : arr) {
            if (str.indexOf(targetId) != -1) {
                return str.split("_")[0];
            }
        }
        return null;
    }

    /**
     * Gets the import check parameters
     *
     * @param field
     * @return
     */
    public static ExcelVerifyEntity getImportVerify(Field field) {
        ExcelVerify verify = field.getAnnotation(ExcelVerify.class);
        if (verify != null) {
            ExcelVerifyEntity entity = new ExcelVerifyEntity();
            entity.setEmail(verify.isEmail());
            entity.setInterHandler(verify.interHandler());
            entity.setMaxLength(verify.maxLength());
            entity.setMinLength(verify.minLength());
            entity.setMobile(verify.isMobile());
            entity.setNotNull(verify.notNull());
            entity.setRegex(verify.regex());
            entity.setRegexTip(verify.regexTip());
            entity.setTel(verify.isTel());
            return entity;
        }
        return null;
    }
}
