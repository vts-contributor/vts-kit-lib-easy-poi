package com.viettel.vtskit.easypoi.poi.excel.entity.params;

import java.util.List;

/**
 * Merge cells using objects
 * 
 */
public class MergeEntity {
	/**
	 * merge start line
	 */
	private int startRow;
	/**
	 * merge end line
	 */
	private int endRow;
	/**
	 * character
	 */
	private String text;
	/**
	 * dependency text
	 */
	private List<String> relyList;

	public MergeEntity() {

	}

	public MergeEntity(String text, int startRow, int endRow) {
		this.text = text;
		this.endRow = endRow;
		this.startRow = startRow;
	}

	public int getEndRow() {
		return endRow;
	}

	public List<String> getRelyList() {
		return relyList;
	}

	public int getStartRow() {
		return startRow;
	}

	public String getText() {
		return text;
	}

	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}

	public void setRelyList(List<String> relyList) {
		this.relyList = relyList;
	}

	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}

	public void setText(String text) {
		this.text = text;
	}
}
