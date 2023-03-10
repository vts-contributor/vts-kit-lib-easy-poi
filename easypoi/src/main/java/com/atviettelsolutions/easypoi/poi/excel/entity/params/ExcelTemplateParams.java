package com.atviettelsolutions.easypoi.poi.excel.entity.params;

import org.apache.poi.ss.usermodel.CellStyle;

import java.io.Serializable;

/**
 * The template convenience is an argument to
 *
 */
public class ExcelTemplateParams implements Serializable {

	/**
     * 
     */
	private static final long serialVersionUID = 1L;
	/**
	 * key
	 */
	private String name;
	/**
	 * Template cell Style
	 */
	private CellStyle cellStyle;
	/**
	 * line height
	 */
	private short height;

	public ExcelTemplateParams() {

	}

	public ExcelTemplateParams(String name, CellStyle cellStyle, short height) {
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

}
