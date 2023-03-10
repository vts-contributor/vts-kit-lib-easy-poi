
package com.atviettelsolutions.easypoi.poi.excel.imports.sax.parse;


import com.atviettelsolutions.easypoi.poi.excel.entity.sax.SaxReadCellEntity;

import java.util.List;

public interface ISaxRowRead {
	/**
	 * Gets the return data
	 * 
	 * @param <T>
	 * @return
	 */
	public <T> List<T> getList();

	/**
	 * Parse the data
	 * 
	 * @param index
	 * @param datas
	 */
	public void parse(int index, List<SaxReadCellEntity> datas);

}
