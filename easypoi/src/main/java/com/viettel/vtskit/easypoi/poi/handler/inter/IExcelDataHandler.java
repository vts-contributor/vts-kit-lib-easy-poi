
package com.viettel.vtskit.easypoi.poi.handler.inter;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.util.Map;
/**
 * Excel Import Export Data Processing Interface
 * 
 * @author caprocute
 */
public interface IExcelDataHandler {

	/**
	 * Export processing method
	 * 
	 * @param obj
	 * @param name
	 * @param value
	 * @return
	 */
	public Object exportHandler(Object obj, String name, Object value);

	/**
	 * Get the fields that need to be processed, import and export are handled uniformly, and reduce the number of fields written
	 * 
	 * @return
	 */
	public String[] getNeedHandlerFields();

	/**
	 * Import processing method Current object, current field name, current value
	 * 
	 * @param obj
	 * @param name
	 * @param value
	 * @return
	 */
	public Object importHandler(Object obj, String name, Object value);

	/**
	 * Set the list of properties that need to be processed
	 * 
	 * @param fields
	 */
	public void setNeedHandlerFields(String[] fields);

	/**
	 * Set up Map import, custom put
	 * 
	 * @param map
	 * @param originKey
	 * @param value
	 */
	public void setMapValue(Map<String, Object> map, String originKey, Object value);
	/**
	 * To get the Hyperlink of this field, version 07 is required and version 03 is not
	 * @param creationHelper
	 * @param obj
	 * @param name
	 * @param value
	 * @return
	 */
	public Hyperlink getHyperlink(CreationHelper creationHelper, Object obj, String name, Object value);

}
