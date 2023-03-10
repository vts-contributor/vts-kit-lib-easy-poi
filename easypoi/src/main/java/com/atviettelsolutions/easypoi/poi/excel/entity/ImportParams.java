package com.atviettelsolutions.easypoi.poi.excel.entity;

import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelVerifyHandler;

import java.util.List;

/**
 * Import parameter settings
 */
public class ImportParams extends ExcelBaseParams {
    /**
     * The number of table header rows, default 0
     */
    private int titleRows = 0;
    /**
     * The number of header rows, default 1
     */
    private int headRows = 1;
    /**
     * The distance between the true value of the field and the column header is 0 by default
     */
    private int startRows = 0;
    /**
     * Primary key setting, how this cell has no value, just skip Or think this is the value below the list
     */
    private int keyIndex = 0;
    /**
     * The position of the sheet to start reading, which defaults to 0
     */
    private int startSheetIndex = 0;

    /**
     * The number of sheets that need to be read to upload a table, 0 by default
     */
    private int sheetNum = 0;
    /**
     * Whether you need to save the uploaded Excel, the default is false
     */
    private boolean needSave = false;
    /**
     * Save the uploaded Excel directory, the default is such as TestEntity, this class save path is
     * upload/excelUpload/Test/yyyyMMddHHmss_*****
     * Save the name upload time_five random numbers
     */
    private String saveUrl = "upload/excelUpload";
    /**
     * Verify the processing interface
     */
    private IExcelVerifyHandler verifyHanlder;
    /**
     * The last number of invalid rows
     */
    private int lastOfInvalidRow = 0;

    /**
     * Headers that do not need to be parsed are only displayed as multiple headers, and no fields are bound to them
     */
    private List<String> ignoreHeaderList;

    /**
     * Picture Columns collection
     */
    private List<String> imageList;

    public int getHeadRows() {
        return headRows;
    }

    public int getKeyIndex() {
        return keyIndex;
    }

    public String getSaveUrl() {
        return saveUrl;
    }

    public int getSheetNum() {
        return sheetNum;
    }

    public int getStartRows() {
        return startRows;
    }

    public int getTitleRows() {
        return titleRows;
    }

    public IExcelVerifyHandler getVerifyHanlder() {
        return verifyHanlder;
    }

    public boolean isNeedSave() {
        return needSave;
    }

    public void setHeadRows(int headRows) {
        this.headRows = headRows;
    }

    public void setKeyIndex(int keyIndex) {
        this.keyIndex = keyIndex;
    }

    public void setNeedSave(boolean needSave) {
        this.needSave = needSave;
    }

    public void setSaveUrl(String saveUrl) {
        this.saveUrl = saveUrl;
    }

    public void setSheetNum(int sheetNum) {
        this.sheetNum = sheetNum;
    }

    public void setStartRows(int startRows) {
        this.startRows = startRows;
    }

    public void setTitleRows(int titleRows) {
        this.titleRows = titleRows;
    }

    public void setVerifyHanlder(IExcelVerifyHandler verifyHanlder) {
        this.verifyHanlder = verifyHanlder;
    }

    public int getLastOfInvalidRow() {
        return lastOfInvalidRow;
    }

    public void setLastOfInvalidRow(int lastOfInvalidRow) {
        this.lastOfInvalidRow = lastOfInvalidRow;
    }

    public List<String> getImageList() {
        return imageList;
    }

    public void setImageList(List<String> imageList) {
        this.imageList = imageList;
    }

    public List<String> getIgnoreHeaderList() {
        return ignoreHeaderList;
    }

    public void setIgnoreHeaderList(List<String> ignoreHeaderList) {
        this.ignoreHeaderList = ignoreHeaderList;
    }

    /**
     * Judge whether to ignore the header based on the text displayed in the header
     *
     * @param text
     * @return
     */
    public boolean isIgnoreHeader(String text) {
        if (ignoreHeaderList != null && ignoreHeaderList.indexOf(text) >= 0) {
            return true;
        }
        return false;
    }

    public int getStartSheetIndex() {
        return startSheetIndex;
    }

    public void setStartSheetIndex(int startSheetIndex) {
        this.startSheetIndex = startSheetIndex;
    }
}
