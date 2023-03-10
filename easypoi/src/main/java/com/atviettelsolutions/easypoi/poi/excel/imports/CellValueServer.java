
package com.atviettelsolutions.easypoi.poi.excel.imports;

import com.atviettelsolutions.easypoi.poi.common.ExcelUtil;
import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelDataHandler;
import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import com.atviettelsolutions.easypoi.poi.exeption.ExcelImportException;
import com.atviettelsolutions.easypoi.poi.exeption.excel.enums.ExcelImportEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelImportEntity;
import com.atviettelsolutions.easypoi.poi.excel.entity.sax.SaxReadCellEntity;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.sql.Time;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Cell value service Judgment typeProcessing data
 * 1. Determine the type in Excel
 * 2. Replace the value according to replace
 * 3. Handler process data
 * 4. Determine the return type Transform data return
 *
 * @author caprocute
 */
public class CellValueServer {

    private static final Logger LOGGER = LoggerFactory.getLogger(CellValueServer.class);

    private List<String> hanlderList = null;

    /**
     * Gets the value within the cell
     *
     * @param xclass
     * @param cell
     * @param entity
     * @return
     */
    private Object getCellValue(String xclass, Cell cell, ExcelImportEntity entity) {
        if (cell == null) {
            return "";
        }
        Object result = null;
        // The date format is special and inconsistent with the cell format
        if ("class java.util.Date".equals(xclass) || ("class java.sql.Time").equals(xclass)) {
            if (CellType.NUMERIC == cell.getCellTypeEnum()) {
                // Date format
                result = cell.getDateCellValue();
            } else {
                cell.setCellType(CellType.STRING);
                result = getDateData(entity, cell.getStringCellValue());
            }
            if (("class java.sql.Time").equals(xclass)) {
                result = new Time(((Date) result).getTime());
            }
        } else if (CellType.NUMERIC == cell.getCellTypeEnum()) {
            result = cell.getNumericCellValue();
        } else if (CellType.BOOLEAN == cell.getCellTypeEnum()) {
            result = cell.getBooleanCellValue();
        } else if (CellType.FORMULA == cell.getCellTypeEnum() && PoiPublicUtil.isNumber(xclass)) {
            //If the cell is an expression and the field is a numeric type
            result = cell.getNumericCellValue();
        } else {
            result = cell.getStringCellValue();
        }
        return result;
    }

    /**
     * Gets date type data
     *
     * @param entity
     * @param value
     * @return
     * @author caprocute
     */
    private Date getDateData(ExcelImportEntity entity, String value) {
        if (StringUtils.isNotEmpty(entity.getFormat()) && StringUtils.isNotEmpty(value)) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            try {
                return format.parse(value);
            } catch (ParseException e) {
                LOGGER.error("Time formatting failed, formatting:{},value:{}", entity.getFormat(), value);
                throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
            }
        }
        return null;
    }

    /**
     * Gets the value of cell
     *
     * @param object
     * @param excelParams
     * @param cell
     * @param titleString
     */
    public Object getValue(IExcelDataHandler dataHanlder, Object object, Cell cell, Map<String,
            ExcelImportEntity> excelParams, String titleString) throws Exception {
        ExcelImportEntity entity = excelParams.get(titleString);
        String xclass = "class java.lang.Object";
        if (!(object instanceof Map)) {
            Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0 ?
                    entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
            Type[] ts = setMethod.getGenericParameterTypes();
            xclass = ts[0].toString();
        }
        Object result = getCellValue(xclass, cell, entity);
        if (entity != null) {
            result = hanlderSuffix(entity.getSuffix(), result);
            result = replaceValue(entity.getReplace(), result, entity.isMultiReplace());
        }
        result = hanlderValue(dataHanlder, object, result, titleString);
        return getValueByType(xclass, result, entity);
    }

    /**
     * Gets the cell value
     *
     * @param dataHanlder
     * @param object
     * @param cellEntity
     * @param excelParams
     * @param titleString
     * @return
     */
    public Object getValue(IExcelDataHandler dataHanlder, Object object, SaxReadCellEntity cellEntity,
                           Map<String, ExcelImportEntity> excelParams, String titleString) {
        ExcelImportEntity entity = excelParams.get(titleString);
        Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0 ?
                entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
        Type[] ts = setMethod.getGenericParameterTypes();
        String xclass = ts[0].toString();
        Object result = cellEntity.getValue();
        result = hanlderSuffix(entity.getSuffix(), result);
        result = replaceValue(entity.getReplace(), result, entity.isMultiReplace());
        result = hanlderValue(dataHanlder, object, result, titleString);
        return getValueByType(xclass, result, entity);
    }

    /**
     * Remove the suffix
     *
     * @param result
     * @param suffix
     * @return
     */
    private Object hanlderSuffix(String suffix, Object result) {
        if (StringUtils.isNotEmpty(suffix) && result != null && result.toString().endsWith(suffix)) {
            String temp = result.toString();
            return temp.substring(0, temp.length() - suffix.length());
        }
        return result;
    }

    /**
     * Gets the return value based on the return type
     *
     * @param xclass
     * @param result
     * @param entity
     * @return
     */
    private Object getValueByType(String xclass, Object result, ExcelImportEntity entity) {
        try {
            if (result == null || "".equals(String.valueOf(result))) {
                return null;
            }
            if ("class java.util.Date".equals(xclass)) {
                return result;
            }
            if ("class java.lang.Boolean".equals(xclass) || "boolean".equals(xclass)) {
                Boolean temp = Boolean.valueOf(String.valueOf(result));
                return temp;

            }
            if ("class java.lang.Double".equals(xclass) || "double".equals(xclass)) {
                Double temp = Double.valueOf(String.valueOf(result));
                return temp;
            }
            if ("class java.lang.Long".equals(xclass) || "long".equals(xclass)) {
                Long temp = Long.valueOf(ExcelUtil.remove0Suffix(String.valueOf(result)));
                return temp;

            }
            if ("class java.lang.Float".equals(xclass) || "float".equals(xclass)) {
                Float temp = Float.valueOf(String.valueOf(result));
                return temp;
            }
            if ("class java.lang.Integer".equals(xclass) || "int".equals(xclass)) {
                Integer temp = Integer.valueOf(ExcelUtil.remove0Suffix(String.valueOf(result)));
                return temp;
            }
            if ("class java.math.BigDecimal".equals(xclass)) {
                BigDecimal temp = new BigDecimal(String.valueOf(result));
                return temp;
            }
            if ("class java.lang.String".equals(xclass)) {
                if (result instanceof String) {
                    return ExcelUtil.remove0Suffix(result);
                }
                if (result instanceof Double) {
                    return PoiPublicUtil.doubleToString((Double) result);
                }
                return ExcelUtil.remove0Suffix(String.valueOf(result));
            }
            return result;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
        }
    }

    /**
     * Call the processing interface to process the value
     *
     * @param dataHanlder
     * @param object
     * @param result
     * @param titleString
     * @return
     */
    private Object hanlderValue(IExcelDataHandler dataHanlder, Object object, Object result, String titleString) {
        if (dataHanlder == null || dataHanlder.getNeedHandlerFields() == null || dataHanlder.getNeedHandlerFields().length == 0) {
            return result;
        }
        if (hanlderList == null) {
            hanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
        }
        if (hanlderList.contains(titleString)) {
            return dataHanlder.importHandler(object, titleString, result);
        }
        return result;
    }


    /**
     * Import supports multi-value substitution
     *
     * @param replace
     * @param result
     * @param multiReplace
     */
    private Object replaceValue(String[] replace, Object result, boolean multiReplace) {
        if (result == null) {
            return "";
        }
        if (replace == null || replace.length <= 0) {
            return result;
        }
        String temp = String.valueOf(result);
        String backValue = "";
        if (temp.indexOf(",") > 0 && multiReplace) {
            //The original value has a comma in it, thinking that he is multi-valued
            String multiReplaces[] = temp.split(",");
            for (String str : multiReplaces) {
                backValue = backValue.concat(replaceSingleValue(replace, str) + ",");
            }
            if (backValue.equals("")) {
                backValue = temp;
            } else {
                backValue = backValue.substring(0, backValue.length() - 1);
            }
        } else {
            backValue = replaceSingleValue(replace, temp);
        }
        if (replace.length > 0 && backValue.equals(temp)) {
            LOGGER.warn("====================Dictionary replacement failed, dictionary value: {}, import value to convert:{}====================", replace, temp);
        }
        return backValue;
    }

    /**
     * Single value substitution , if not found, the original value is returned
     */
    private String replaceSingleValue(String[] replace, String temp) {
        String[] tempArr;
        for (int i = 0; i < replace.length; i++) {
            tempArr = getValueArr(replace[i]);
            if (temp.equals(tempArr[0]) || temp.replace("_", "---").equals(tempArr[0])) {
                if (tempArr[1].contains("---")) {
                    return tempArr[1].replace("---", "_");
                }
                return tempArr[1];
            }
        }
        return temp;
    }

    /**
     * The dictionary text contains multiple underscores, take the last one (to solve the null case)
     *
     * @param val
     * @return
     */
    public String[] getValueArr(String val) {
        int i = val.lastIndexOf("_");//The position of the last separator
        String[] c = new String[2];
        c[0] = val.substring(0, i); //label
        c[1] = val.substring(i + 1); //key
        return c;
    }

}
