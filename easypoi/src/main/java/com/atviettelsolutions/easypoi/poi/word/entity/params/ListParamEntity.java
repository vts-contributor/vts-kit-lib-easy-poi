
package com.atviettelsolutions.easypoi.poi.word.entity.params;

/**
 * Excel Object export structure
 * 
 */
public class ListParamEntity {
	// Unique values, reused in traversals
	public static final String SINGLE = "single";
	// Belongs to an array type
	public static final String LIST = "list";
	/**
	 * Property name
	 */
	private String name;
	/**
	 * target
	 */
	private String target;
	/**
	 * When it is the only value, the value is evaluated directly
	 */
	private Object value;
	/**
	 * data type,SINGLE || LIST
	 */
	private String type;

	public ListParamEntity() {

	}

	public ListParamEntity(String name, Object value) {
		this.name = name;
		this.value = value;
		this.type = LIST;
	}

	public ListParamEntity(String name, String target) {
		this.name = name;
		this.target = target;
		this.type = LIST;
	}

	public String getName() {
		return name;
	}

	public String getTarget() {
		return target;
	}

	public String getType() {
		return type;
	}

	public Object getValue() {
		return value;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setTarget(String target) {
		this.target = target;
	}

	public void setType(String type) {
		this.type = type;
	}

	public void setValue(Object value) {
		this.value = value;
	}
}
