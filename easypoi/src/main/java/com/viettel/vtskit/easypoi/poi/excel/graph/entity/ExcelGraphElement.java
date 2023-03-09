/**
 * 
 */
package com.viettel.vtskit.easypoi.poi.excel.graph.entity;


import com.viettel.vtskit.easypoi.poi.excel.graph.constant.ExcelGraphElementType;

public class ExcelGraphElement
{
	private Integer startRowNum;
	private Integer endRowNum;
	private Integer startColNum;
	private Integer endColNum;
	private Integer elementType= ExcelGraphElementType.STRING_TYPE;
	
	
	public Integer getStartRowNum()
	{
		return startRowNum;
	}
	public void setStartRowNum(Integer startRowNum)
	{
		this.startRowNum = startRowNum;
	}
	public Integer getEndRowNum()
	{
		return endRowNum;
	}
	public void setEndRowNum(Integer endRowNum)
	{
		this.endRowNum = endRowNum;
	}
	public Integer getStartColNum()
	{
		return startColNum;
	}
	public void setStartColNum(Integer startColNum)
	{
		this.startColNum = startColNum;
	}
	public Integer getEndColNum()
	{
		return endColNum;
	}
	public void setEndColNum(Integer endColNum)
	{
		this.endColNum = endColNum;
	}
	public Integer getElementType()
	{
		return elementType;
	}
	public void setElementType(Integer elementType)
	{
		this.elementType = elementType;
	}
}
