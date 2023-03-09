package com.viettel.vtskit.easypoi.poi.excel.entity.params;

import java.util.List;

/**
 * Excel imports tool classes to map cell types
 */
public class ExcelImportEntity extends ExcelBaseEntity {
    /**
     * Corresponding Collection NAME
     */
    private String collectionName;
    /**
     * Address to save pictures When saveType is set to 34, this value can be set to: local, minio, alioss
     */
    private String saveUrl;
    /**
     * The type of saved picture, 1 is file_old, 2 is database byte, 3 file address_new, 4 network address
     */
    private int saveType;
    /**
     * Corresponding to export Type
     */
    private String classType;
    /**
     * checkParameter
     */
    private ExcelVerifyEntity verify;
    /**
     * suffix
     */
    private String suffix;

    private List<ExcelImportEntity> list;

    public String getClassType() {
        return classType;
    }

    public String getCollectionName() {
        return collectionName;
    }

    public List<ExcelImportEntity> getList() {
        return list;
    }

    public int getSaveType() {
        return saveType;
    }

    public String getSaveUrl() {
        return saveUrl;
    }

    public ExcelVerifyEntity getVerify() {
        return verify;
    }

    public void setClassType(String classType) {
        this.classType = classType;
    }

    public void setCollectionName(String collectionName) {
        this.collectionName = collectionName;
    }

    public void setList(List<ExcelImportEntity> list) {
        this.list = list;
    }

    public void setSaveType(int saveType) {
        this.saveType = saveType;
    }

    public void setSaveUrl(String saveUrl) {
        this.saveUrl = saveUrl;
    }

    public void setVerify(ExcelVerifyEntity verify) {
        this.verify = verify;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

}
