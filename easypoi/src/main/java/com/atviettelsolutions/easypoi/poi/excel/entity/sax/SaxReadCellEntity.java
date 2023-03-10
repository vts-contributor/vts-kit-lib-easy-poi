package com.atviettelsolutions.easypoi.poi.excel.entity.sax;

import com.atviettelsolutions.easypoi.poi.excel.entity.enmus.CellValueType;

/**
 * Cell object
 */
public class SaxReadCellEntity {
    /**
     * value type
     */
    private CellValueType cellType;
    /**
     * value
     */
    private Object value;

    public SaxReadCellEntity(CellValueType cellType, Object value) {
        this.cellType = cellType;
        this.value = value;
    }

    public CellValueType getCellType() {
        return cellType;
    }

    public void setCellType(CellValueType cellType) {
        this.cellType = cellType;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return "[type=" + cellType.toString() + ",value=" + value + "]";
    }

}
