package com.viettel.vtskit.easypoi.poi.excel.entity.params;

import org.apache.poi.ss.usermodel.CellStyle;

import java.io.Serializable;
import java.util.Stack;

/**
 * The template for each is a parameter of
 */
public class ExcelForEachParams implements Serializable {

    /**
     *
     */
    private static final long serialVersionUID = 1L;
    /**
     * key
     */
    private String name;
    /**
     * key
     */
    private Stack<String> tempName;
    /**
     * Template cell Style
     */
    private CellStyle cellStyle;
    /**
     * 行高
     */
    private short height;
    /**
     * constant value
     */
    private String constValue;
    /**
     * column merge
     */
    private int colspan = 1;
    /**
     * LINE-MERGE
     */
    private int rowspan = 1;
    /**
     * line merge
     */
    private boolean collectCell;

    public ExcelForEachParams() {

    }

    public ExcelForEachParams(String name, CellStyle cellStyle, short height) {
        this.name = name;
        this.cellStyle = cellStyle;
        this.height = height;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public short getHeight() {
        return height;
    }

    public void setHeight(short height) {
        this.height = height;
    }

    public String getConstValue() {
        return constValue;
    }

    public void setConstValue(String constValue) {
        this.constValue = constValue;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public boolean isCollectCell() {
        return collectCell;
    }

    public void setCollectCell(boolean collectCell) {
        this.collectCell = collectCell;
    }

    public Stack<String> getTempName() {
        return tempName;
    }

    public void setTempName(Stack<String> tempName) {
        this.tempName = tempName;
    }
}
