package com.atviettelsolutions.easypoi.poi.exeption.word.enmus;
public enum WordExportEnum {

	EXCEL_PARAMS_ERROR("Excel export parameter error"), EXCEL_HEAD_HAVA_NULL("Some fields in Excel table header are empty"), EXCEL_NO_HEAD("Excel has no header");

	private String msg;

	WordExportEnum(String msg) {
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

}
