package com.viettel.vtskit.easypoi.poi.excel.entity.params;

import java.util.Map;

/**
 * Excel's Collection for
 *
 * @author caprocute
 * @version 1.0
 * @date 2013-9-26
 */
public class ExcelCollectionParams {

    /**
     * The name corresponding to the collection
     */
    private String name;
    /**
     * Excel column names
     */
    private String excelName;
    /**
     * entity object
     */
    private Class<?> type;
    /**
     * The parameter set entity object under this list
     */
    private Map<String, ExcelImportEntity> excelParams;

    public Map<String, ExcelImportEntity> getExcelParams() {
        return excelParams;
    }

    public String getName() {
        return name;
    }

    public Class<?> getType() {
        return type;
    }

    public void setExcelParams(Map<String, ExcelImportEntity> excelParams) {
        this.excelParams = excelParams;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setType(Class<?> type) {
        this.type = type;
    }

    public String getExcelName() {
        return excelName;
    }

    public void setExcelName(String excelName) {
        this.excelName = excelName;
    }
}
