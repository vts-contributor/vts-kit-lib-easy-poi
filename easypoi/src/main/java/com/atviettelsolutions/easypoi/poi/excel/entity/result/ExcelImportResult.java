package com.atviettelsolutions.easypoi.poi.excel.entity.result;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * import return class
 *
 */
public class ExcelImportResult<T> {

	/**
	 * result set
	 */
	private List<T> list;

	/**
	 * Whether there is a verification failure
	 */
	private boolean verfiyFail;

	/**
	 * data source
	 */
	private Workbook workbook;

	public ExcelImportResult() {

	}

	public ExcelImportResult(List<T> list, boolean verfiyFail, Workbook workbook) {
		this.list = list;
		this.verfiyFail = verfiyFail;
		this.workbook = workbook;
	}

	public List<T> getList() {
		return list;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public boolean isVerfiyFail() {
		return verfiyFail;
	}

	public void setList(List<T> list) {
		this.list = list;
	}

	public void setVerfiyFail(boolean verfiyFail) {
		this.verfiyFail = verfiyFail;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

}
