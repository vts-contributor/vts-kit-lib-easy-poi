
package com.atviettelsolutions.easypoi.poi.handler.impl;

import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelDataHandler;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.util.Map;
/**
 * Data processing is implemented by default and returns empty
 * 
 * @author caprocute
 */
public abstract class ExcelDataHandlerDefaultImpl implements IExcelDataHandler {
	/**
	 * The fields that need to be processed
	 */
	private String[] needHandlerFields;

	@Override
	public Object exportHandler(Object obj, String name, Object value) {
		return value;
	}

	@Override
	public String[] getNeedHandlerFields() {
		return needHandlerFields;
	}

	@Override
	public Object importHandler(Object obj, String name, Object value) {
		return value;
	}

	@Override
	public void setNeedHandlerFields(String[] needHandlerFields) {
		this.needHandlerFields = needHandlerFields;
	}

	@Override
	public void setMapValue(Map<String, Object> map, String originKey, Object value) {
		map.put(originKey, value);
	}

	@Override
	public Hyperlink getHyperlink(CreationHelper creationHelper, Object obj, String name, Object value) {
		return null;
	}
}
