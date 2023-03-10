package com.atviettelsolutions.easypoi.poi.excel.entity;

import com.atviettelsolutions.easypoi.poi.excel.entity.enmus.ExcelType;
import com.atviettelsolutions.easypoi.poi.excel.export.styler.ExcelExportStylerDefaultImpl;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Excel export parameters
 *
 */
public class ExportParams extends ExcelBaseParams {

	/**
	 * form name
	 */
	private String title;

	/**
	 * form name
	 */
	private short titleHeight = 10;

	/**
	 * second row name
	 */
	private String secondTitle;

	/**
	 * form name
	 */
	private short secondTitleHeight = 8;
	/**
	 * sheetName
	 */
	private String sheetName;
	/**
	 * The properties of the filter
	 */
	private String[] exclusions;
	/**
	 * Whether to add the need to need
	 */
	private boolean addIndex;
	/**
	 * Whether to add the need to need
	 */
	private String indexName = "serial number";
	/**
	 * frozen column
	 */
	private int freezeCol;
	/**
	 * header color
	 */
	private short color = IndexedColors.WHITE.index;
	/**
	 * The color of the property description row For example: HSSFColor.SKY_BLUE.index default
	 */
	private short headerColor = IndexedColors.SKY_BLUE.index;
	/**
	 * Excel export version
	 */
	private ExcelType type = ExcelType.HSSF;
	/**
	 * Excel exports style
	 */
	private Class<?> style = ExcelExportStylerDefaultImpl.class;
	/**
	 * Whether to create a table header
	 */
	private boolean isCreateHeadRows = true;

	/**
	 * Local file storage root path base path
	 */
	private String imageBasePath;
	/**
	 * Whether to fix the header
	 */
	private boolean isFixedTitle     = true;
	/**
	* Single sheet maximum version 03 default 6W rows, 07 default 100W
	 */
	private int     maxNum           = 0;
	/**
	 * When exporting the height of each column in excel in characters, one kanji=2 characters global setting, take precedence
	 */
	private short height = 0;

	/**
	 * read only
	 */
	private boolean readonly = false;
	public ExportParams() {

	}

	public ExportParams(String title, String sheetName) {
		this.title = title;
		this.sheetName = sheetName;
	}

	public ExportParams(String title, String sheetName, ExcelType type) {
		this.title = title;
		this.sheetName = sheetName;
		this.type = type;
	}

	public ExportParams(String title, String secondTitle, String sheetName) {
		this.title = title;
		this.secondTitle = secondTitle;
		this.sheetName = sheetName;
	}

	public ExportParams(String title, String secondTitle, String sheetName,String imageBasePath) {
		this.title = title;
		this.secondTitle = secondTitle;
		this.sheetName = sheetName;
		this.imageBasePath = imageBasePath;
	}

	public short getColor() {
		return color;
	}

	public String[] getExclusions() {
		return exclusions;
	}

	public short getHeaderColor() {
		return headerColor;
	}

	public String getSecondTitle() {
		return secondTitle;
	}

	public short getSecondTitleHeight() {
		return (short) (secondTitleHeight * 50);
	}

	public String getSheetName() {
		return sheetName;
	}

	public String getTitle() {
		return title;
	}

	public short getTitleHeight() {
		return (short) (titleHeight * 50);
	}

	public boolean isAddIndex() {
		return addIndex;
	}

	public void setAddIndex(boolean addIndex) {
		this.addIndex = addIndex;
	}

	public void setColor(short color) {
		this.color = color;
	}

	public void setExclusions(String[] exclusions) {
		this.exclusions = exclusions;
	}

	public void setHeaderColor(short headerColor) {
		this.headerColor = headerColor;
	}

	public void setSecondTitle(String secondTitle) {
		this.secondTitle = secondTitle;
	}

	public void setSecondTitleHeight(short secondTitleHeight) {
		this.secondTitleHeight = secondTitleHeight;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public void setTitleHeight(short titleHeight) {
		this.titleHeight = titleHeight;
	}

	public ExcelType getType() {
		return type;
	}

	public void setType(ExcelType type) {
		this.type = type;
	}

	public String getIndexName() {
		return indexName;
	}

	public void setIndexName(String indexName) {
		this.indexName = indexName;
	}

	public Class<?> getStyle() {
		return style;
	}

	public void setStyle(Class<?> style) {
		this.style = style;
	}

	public int getFreezeCol() {
		return freezeCol;
	}

	public void setFreezeCol(int freezeCol) {
		this.freezeCol = freezeCol;
	}

	public boolean isCreateHeadRows() {
		return isCreateHeadRows;
	}

	public void setCreateHeadRows(boolean isCreateHeadRows) {
		this.isCreateHeadRows = isCreateHeadRows;
	}

	public String getImageBasePath() {
		return imageBasePath;
	}

	public void setImageBasePath(String imageBasePath) {
		this.imageBasePath = imageBasePath;
	}

	public int getMaxNum() {
		return maxNum;
	}

	public void setMaxNum(int maxNum) {
		this.maxNum = maxNum;
	}

	public short getHeight() {
		return height == -1 ? -1 : (short) (height * 50);
	}

	public void setHeight(short height) {
		this.height = height;
	}

	public boolean isFixedTitle() {
		return isFixedTitle;
	}

	public void setFixedTitle(boolean fixedTitle) {
		isFixedTitle = fixedTitle;
	}

	public boolean isReadonly() {
		return readonly;
	}

	public void setReadonly(boolean readonly) {
		this.readonly = readonly;
	}
}
