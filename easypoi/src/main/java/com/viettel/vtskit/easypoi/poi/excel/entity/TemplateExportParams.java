package com.viettel.vtskit.easypoi.poi.excel.entity;


import com.viettel.vtskit.easypoi.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * Template export parameter settings
 *
 * @author caprocute
 * @version 1.0
 */
public class TemplateExportParams extends ExcelBaseParams {

    /**
     * Output all sheets
     */
    private boolean scanAllsheet = false;
    /**
     * The path to the template
     */
    private String templateUrl;

    /**
     * The first sheetNum that needs to be exported, the default is the 0th
     */
    private Integer[] sheetNum = new Integer[]{0};

    /**
     * This sheetName uses the original without filling in
     */
    private String[] sheetName;

    /**
     * The number of rows in the table column header, default 1
     */
    private int headingRows = 1;

    /**
     * Table column header starting row, default 1
     */
    private int headingStartRow = 1;
    /**
     * Sets the NUM of the data source
     */
    private int dataSheetNum = 0;
    /**
     * Excel exports style
     */
    private Class<?> style = ExcelExportStylerDefaultImpl.class;
    /**
     * The local variable used by FOR EACH
     */
    private String tempParams = "t";
    //Column loop support
    private boolean colForEach = false;

    /**
     * Default constructor
     */
    public TemplateExportParams() {

    }

    /**
     * Constructor
     *
     * @param templateUrl
     * @param scanAllsheet
     * @param sheetName
     */
    public TemplateExportParams(String templateUrl, boolean scanAllsheet, String... sheetName) {
        this.templateUrl = templateUrl;
        this.scanAllsheet = scanAllsheet;
        if (sheetName != null && sheetName.length > 0) {
            this.sheetName = sheetName;

        }
    }

    /**
     * Constructor
     *
     * @param templateUrl
     * @param sheetNum
     */
    public TemplateExportParams(String templateUrl, Integer... sheetNum) {
        this.templateUrl = templateUrl;
        if (sheetNum != null && sheetNum.length > 0) {
            this.sheetNum = sheetNum;
        }
    }

    /**
     * A single sheet output constructor
     *
     * @param templateUrl
     * @param sheetName
     * @param sheetNum
     */
    public TemplateExportParams(String templateUrl, String sheetName, Integer... sheetNum) {
        this.templateUrl = templateUrl;
        this.sheetName = new String[]{sheetName};
        if (sheetNum != null && sheetNum.length > 0) {
            this.sheetNum = sheetNum;
        }
    }

    public int getHeadingRows() {
        return headingRows;
    }

    public int getHeadingStartRow() {
        return headingStartRow;
    }

    public String[] getSheetName() {
        return sheetName;
    }

    public Integer[] getSheetNum() {
        return sheetNum;
    }

    public String getTemplateUrl() {
        return templateUrl;
    }

    public void setHeadingRows(int headingRows) {
        this.headingRows = headingRows;
    }

    public void setHeadingStartRow(int headingStartRow) {
        this.headingStartRow = headingStartRow;
    }

    public void setSheetName(String[] sheetName) {
        this.sheetName = sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = new String[]{sheetName};
    }

    public void setSheetNum(Integer[] sheetNum) {
        this.sheetNum = sheetNum;
    }

    public void setSheetNum(Integer sheetNum) {
        this.sheetNum = new Integer[]{sheetNum};
    }

    public void setTemplateUrl(String templateUrl) {
        this.templateUrl = templateUrl;
    }

    public Class<?> getStyle() {
        return style;
    }

    public void setStyle(Class<?> style) {
        this.style = style;
    }

    public int getDataSheetNum() {
        return dataSheetNum;
    }

    public void setDataSheetNum(int dataSheetNum) {
        this.dataSheetNum = dataSheetNum;
    }

    public boolean isScanAllsheet() {
        return scanAllsheet;
    }

    public void setScanAllsheet(boolean scanAllsheet) {
        this.scanAllsheet = scanAllsheet;
    }

    public String getTempParams() {
        return tempParams;
    }

    public void setTempParams(String tempParams) {
        this.tempParams = tempParams;
    }

    public boolean isColForEach() {
        return colForEach;
    }

    public void setColForEach(boolean colForEach) {
        this.colForEach = colForEach;
    }
}
