package com.viettel.vtskit.easypoi.poi.excel.entity.result;

/**
 * Excel import processing return results
 *
 */
public class ExcelVerifyHanlderResult {
	/**
	 * is it right or not
	 */
	private boolean success;
	/**
	 * error message
	 */
	private String msg;

	public ExcelVerifyHanlderResult() {

	}

	public ExcelVerifyHanlderResult(boolean success) {
		this.success = success;
	}

	public ExcelVerifyHanlderResult(boolean success, String msg) {
		this.success = success;
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public boolean isSuccess() {
		return success;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

	public void setSuccess(boolean success) {
		this.success = success;
	}

}
