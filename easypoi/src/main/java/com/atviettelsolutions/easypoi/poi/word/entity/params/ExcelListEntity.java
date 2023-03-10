
package com.atviettelsolutions.easypoi.poi.word.entity.params;

import com.atviettelsolutions.easypoi.poi.excel.entity.ExcelBaseParams;
import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelDataHandler;

import java.util.List;

/**
 * Excel Export the object
 * 
 * @author caprocute
 */
public class ExcelListEntity extends ExcelBaseParams {

	/**
	 * Data source
	 */
	private List<?> list;

	/**
	 * Entity class object
	 */
	private Class<?> clazz;

	/**
	 * The number of header rows
	 */
	private int headRows = 1;

	public ExcelListEntity() {

	}

	public ExcelListEntity(List<?> list, Class<?> clazz) {
		this.list = list;
		this.clazz = clazz;
	}

	public ExcelListEntity(List<?> list, Class<?> clazz, IExcelDataHandler dataHanlder) {
		this.list = list;
		this.clazz = clazz;
		setDataHanlder(dataHanlder);
	}

	public ExcelListEntity(List<?> list, Class<?> clazz, IExcelDataHandler dataHanlder, int headRows) {
		this.list = list;
		this.clazz = clazz;
		this.headRows = headRows;
		setDataHanlder(dataHanlder);
	}

	public ExcelListEntity(List<?> list, Class<?> clazz, int headRows) {
		this.list = list;
		this.clazz = clazz;
		this.headRows = headRows;
	}

	public Class<?> getClazz() {
		return clazz;
	}

	public int getHeadRows() {
		return headRows;
	}

	public List<?> getList() {
		return list;
	}

	public void setClazz(Class<?> clazz) {
		this.clazz = clazz;
	}

	public void setHeadRows(int headRows) {
		this.headRows = headRows;
	}

	public void setList(List<?> list) {
		this.list = list;
	}

}
