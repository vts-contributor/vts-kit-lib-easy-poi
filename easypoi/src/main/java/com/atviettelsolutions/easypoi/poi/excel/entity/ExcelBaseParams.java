package com.atviettelsolutions.easypoi.poi.excel.entity;

import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelDataHandler;

/**
 * Basic parameters
 *
 */
public class ExcelBaseParams {

	/**
	 * The data processing interface is mainly based on this, replace and format are all behind this
	 */
	private IExcelDataHandler dataHanlder;

	public IExcelDataHandler getDataHanlder() {
		return dataHanlder;
	}

	public void setDataHanlder(IExcelDataHandler dataHanlder) {
		this.dataHanlder = dataHanlder;
	}

}
