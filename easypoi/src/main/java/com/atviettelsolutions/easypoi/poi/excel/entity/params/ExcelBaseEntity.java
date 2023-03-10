package com.atviettelsolutions.easypoi.poi.excel.entity.params;

import java.lang.reflect.Method;
import java.util.List;

/**
 * Excel import and export base object class
 */
public class ExcelBaseEntity {
    /**
     * corresponding name
     */
    protected String name;
    /**
     * corresponding type
     */
    private int type = 1;
    /**
     * database format
     */
    private String databaseFormat;
    /**
     * export date format
     */
    private String format;

    /**
     * Number formatting, the parameter is Pattern, and the object used is Decimal Format
     */
    private String numFormat;
    /**
     * Replacement value expression: "male_1","female_0"
     */
    private String[] replace;
    /**
     * Whether to replace multiple values
     */
    private boolean multiReplace;

    /**
     * header group name
     */
    private String groupName;

    /**
     * set/get method
     */
    private Method method;
    /**
     * fixed column
     */
    private Integer fixedIndex;
    /**
     * dictionary name
     */
    private String dict;
    /**
     * Is this a hyperlink? If it needs to implement the interface to return the object
     */
    private boolean hyperlink;

    private List<Method> methods;

    public String getDatabaseFormat() {
        return databaseFormat;
    }

    public String getFormat() {
        return format;
    }

    public Method getMethod() {
        return method;
    }

    public List<Method> getMethods() {
        return methods;
    }

    public String getName() {
        return name;
    }

    public String[] getReplace() {
        return replace;
    }

    public int getType() {
        return type;
    }

    public void setDatabaseFormat(String databaseFormat) {
        this.databaseFormat = databaseFormat;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public void setMethod(Method method) {
        this.method = method;
    }

    public void setMethods(List<Method> methods) {
        this.methods = methods;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setReplace(String[] replace) {
        this.replace = replace;
    }

    public void setType(int type) {
        this.type = type;
    }

    public boolean isMultiReplace() {
        return multiReplace;
    }

    public void setMultiReplace(boolean multiReplace) {
        this.multiReplace = multiReplace;
    }

    public String getNumFormat() {
        return numFormat;
    }

    public void setNumFormat(String numFormat) {
        this.numFormat = numFormat;
    }

    public String getGroupName() {
        return groupName;
    }

    public void setGroupName(String groupName) {
        this.groupName = groupName;
    }

    public Integer getFixedIndex() {
        return fixedIndex;
    }

    public void setFixedIndex(Integer fixedIndex) {
        this.fixedIndex = fixedIndex;
    }

    public String getDict() {
        return dict;
    }

    public void setDict(String dict) {
        this.dict = dict;
    }

    public boolean isHyperlink() {
        return hyperlink;
    }

    public void setHyperlink(boolean hyperlink) {
        this.hyperlink = hyperlink;
    }
}
