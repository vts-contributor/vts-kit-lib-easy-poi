package com.atviettelsolutions.easypoi.poi.excel.entity.enmus;


import com.atviettelsolutions.easypoi.poi.excel.export.styler.ExcelExportStylerBorderImpl;
import com.atviettelsolutions.easypoi.poi.excel.export.styler.ExcelExportStylerColorImpl;
import com.atviettelsolutions.easypoi.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * Several default styles provided by the plugin
 *
 */
public enum ExcelStyleType {

	NONE("default style", ExcelExportStylerDefaultImpl.class),
	BORDER("border style", ExcelExportStylerBorderImpl.class),
	COLOR("spaced line style", ExcelExportStylerColorImpl.class);

	private String name;
	private Class<?> clazz;

	ExcelStyleType(String name, Class<?> clazz) {
		this.name = name;
		this.clazz = clazz;
	}

	public Class<?> getClazz() {
		return clazz;
	}

	public String getName() {
		return name;
	}

}
