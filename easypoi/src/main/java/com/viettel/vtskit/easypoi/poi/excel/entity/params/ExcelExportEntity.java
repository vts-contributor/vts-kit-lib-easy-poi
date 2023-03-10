package com.viettel.vtskit.easypoi.poi.excel.entity.params;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel export tool class, to map the cell type
 * 
 */
public class ExcelExportEntity extends ExcelBaseEntity implements Comparable<ExcelExportEntity> {

	/**
	 * If it is a MAP export, this is the key of the map
	 */
	private Object key;

	private double width = 10;

	private double height = 10;

	/**
	 * The type of picture, 1 is the file address (class directory),
	 * 2 is the database byte, 3 is the file address (disk directory), 4 is the network picture
	 */
	private int exportImageType = 3;

	/**
	 * The image storage location (disk directory) is used to export and obtain the absolute path of the image
	 */
	private String imageBasePath;

	/**
	 * sort order
	 */
	private int orderNum = 0;

	/**
	 * Whether to support newline
	 */
	private boolean isWrap;

	/**
	 * Do you need to merge
	 */
	private boolean needMerge;
	/**
	 * Merge cells vertically
	 */
	private boolean mergeVertical;
	/**
	 * merge dependencies
	 */
	private int[] mergeRely;
	/**
	 * suffix
	 */
	private String suffix;
	/**
	 * statistics
	 */
	private boolean isStatistics;

	/**
	 * whether to merge horizontally
	 */
	private boolean colspan;

	/**
	 * column names to be merged horizontally
	 */
	private List<String> subColumnList;

	/**
	 * The name of the parent header
	 */
	private String groupName;
	/**
	 *  Whether to hide the column
	 */
	private boolean isColumnHidden;

	private List<ExcelExportEntity> list;

	public ExcelExportEntity() {

	}

	public ExcelExportEntity(String name) {
		super.name = name;
	}

	public ExcelExportEntity(String name, Object key) {
		super.name = name;
		this.key = key;
	}

	/**
	 * Constructor
	 * @param name descriptionText
	 * @param key Storage key If it is exported by MAP, this is the key of the map
	 * @param colspan Whether it is a merged column (columns a and b share a header c, then a, b, and c all need to be set to true)
	 */
	public ExcelExportEntity(String name, Object key, boolean colspan) {
		super.name = name;
		this.key = key;
		this.colspan = colspan;
		this.needMerge = colspan;
	}

	public ExcelExportEntity(String name, Object key, int width) {
		super.name = name;
		this.width = width;
		this.key = key;
	}

	public int getExportImageType() {
		return exportImageType;
	}

	public double getHeight() {
		return height;
	}

	public Object getKey() {
		return key;
	}

	public List<ExcelExportEntity> getList() {
		return list;
	}

	public int[] getMergeRely() {
		return mergeRely == null ? new int[0] : mergeRely;
	}

	public int getOrderNum() {
		return orderNum;
	}

	public double getWidth() {
		return width;
	}

	public boolean isMergeVertical() {
		return mergeVertical;
	}

	public boolean isNeedMerge() {
		return needMerge;
	}

	public boolean isWrap() {
		return isWrap;
	}

	public void setExportImageType(int exportImageType) {
		this.exportImageType = exportImageType;
	}

	public void setHeight(double height) {
		this.height = height;
	}

	public void setKey(Object key) {
		this.key = key;
	}

	public void setList(List<ExcelExportEntity> list) {
		this.list = list;
	}

	public void setMergeRely(int[] mergeRely) {
		this.mergeRely = mergeRely;
	}

	public void setMergeVertical(boolean mergeVertical) {
		this.mergeVertical = mergeVertical;
	}

	public void setNeedMerge(boolean needMerge) {
		this.needMerge = needMerge;
	}

	public void setOrderNum(int orderNum) {
		this.orderNum = orderNum;
	}

	public void setWidth(double width) {
		this.width = width;
	}

	public void setWrap(boolean isWrap) {
		this.isWrap = isWrap;
	}

	public String getSuffix() {
		return suffix;
	}

	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}

	public boolean isStatistics() {
		return isStatistics;
	}

	public void setStatistics(boolean isStatistics) {
		this.isStatistics = isStatistics;
	}

	public String getImageBasePath() {
		return imageBasePath;
	}

	public void setImageBasePath(String imageBasePath) {
		this.imageBasePath = imageBasePath;
	}

	public boolean isColspan() {
		return colspan;
	}

	public void setColspan(boolean colspan) {
		this.colspan = colspan;
	}

	public List<String> getSubColumnList() {
		return subColumnList;
	}

	public void setSubColumnList(List<String> subColumnList) {
		this.subColumnList = subColumnList;
	}

	public String getGroupName() {
		return groupName;
	}

	public void setGroupName(String groupName) {
		this.groupName = groupName;
	}

	public boolean isColumnHidden() {
		return isColumnHidden;
	}

	public void setColumnHidden(boolean columnHidden) {
		isColumnHidden = columnHidden;
	}

	/**
	 * Whether it is a merged sub-column
	 * @return
	 */
	public boolean isSubColumn(){
		return this.colspan && (this.subColumnList==null || this.subColumnList.size()==0);
	}

	/**
	 * Whether to merge parent columns
	 * @return
	 */
	public boolean isMergeColumn(){
		return this.colspan && this.subColumnList!=null && this.subColumnList.size()>0;
	}


	/**
	 * Get merged sub-columns
	 * @param all
	 * @return
	 */
	public List<ExcelExportEntity> initSubExportEntity(List<ExcelExportEntity> all){
		List<ExcelExportEntity> sub = new ArrayList<ExcelExportEntity>();
		for (ExcelExportEntity temp : all) {
			if(this.subColumnList.contains(temp.getKey())){
				sub.add(temp);
			}
		}
		this.setList(sub);
		return sub;
	}

	@Override
	public int compareTo(ExcelExportEntity prev) {
		return this.getOrderNum() - prev.getOrderNum();
	}

}
