
package com.atviettelsolutions.easypoi.poi.word.entity;

/**
 * Word export, picture settings and picture information
 * 
 * @author caprocute
 * @version 1.0
 */
public class WordImageEntity {

	public static String URL = "url";
	public static String Data = "data";
	/**
	 * Image input method
	 */
	private String type = URL;
	/**
	 * Image width
	 */
	private int width;
	// Image height
	private int height;
	// Image address
	private String url;
	// Picture information
	private byte[] data;

	public WordImageEntity() {

	}

	public WordImageEntity(byte[] data, int width, int height) {
		this.data = data;
		this.width = width;
		this.height = height;
		this.type = Data;
	}

	public WordImageEntity(String url, int width, int height) {
		this.url = url;
		this.width = width;
		this.height = height;
	}

	public byte[] getData() {
		return data;
	}

	public int getHeight() {
		return height;
	}

	public String getType() {
		return type;
	}

	public String getUrl() {
		return url;
	}

	public int getWidth() {
		return width;
	}

	public void setData(byte[] data) {
		this.data = data;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public void setType(String type) {
		this.type = type;
	}

	public void setUrl(String url) {
		this.url = url;
	}

	public void setWidth(int width) {
		this.width = width;
	}

}
