/**
 * 
 */
package com.atviettelsolutions.easypoi.poi.excel.graph.entity;

import java.util.List;


public interface ExcelGraph
{
	public ExcelGraphElement getCategory();
	public List<ExcelGraphElement> getValueList();
	public Integer getGraphType();
	public List<ExcelTitleCell> getTitleCell();
	public List<String> getTitle();
}
