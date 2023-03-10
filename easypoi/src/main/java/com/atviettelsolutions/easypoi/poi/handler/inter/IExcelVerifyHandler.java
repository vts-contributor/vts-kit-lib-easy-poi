
package com.atviettelsolutions.easypoi.poi.handler.inter;


import com.atviettelsolutions.easypoi.poi.excel.entity.result.ExcelVerifyHanlderResult;

/**
 * Import the validation interface
 * 
 * @author caprocute
 */
public interface IExcelVerifyHandler {

	/**
	 * Get the fields that need to be processed, import and export are handled uniformly, and reduce the number of fields written
	 * 
	 * @return
	 */
	public String[] getNeedVerifyFields();

	/**
	 * Get the fields that need to be processed, import and export are handled uniformly, and reduce the number of fields written
	 * 
	 * @return
	 */
	public void setNeedVerifyFields(String[] arr);

	/**
	 * Export processing method
	 * 
	 * @param obj
	 *
	 * @param name
	 *
	 * @param value
	 *
	 * @return
	 */
	public ExcelVerifyHanlderResult verifyHandler(Object obj, String name, Object value);

}
