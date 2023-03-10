package com.atviettelsolutions.easypoi.poi.exeption.excel.enums;

/**
 * Export exception type enumeration
 *
 */
public enum ExcelExportEnum {

	PARAMETER_ERROR("Excel export parameter error"), EXPORT_ERROR("Excel export error"),TEMPLATE_ERROR ("Excel template error");;

	private String msg;

	ExcelExportEnum(String msg) {
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

}
