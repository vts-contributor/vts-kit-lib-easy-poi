package com.viettel.vtskit.easypoi.poi.exeption.excel.enums;

public enum ExcelImportEnum {

	GET_VALUE_ERROR("Excel Failed to get value"), VERIFY_ERROR("Value validation failed");

	private String msg;

	ExcelImportEnum(String msg) {
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

}
