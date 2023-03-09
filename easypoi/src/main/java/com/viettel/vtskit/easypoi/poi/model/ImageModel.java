package com.viettel.vtskit.easypoi.poi.model;

/**
 * @author hieuhk1 - 09/03/2023
 **/
public class ImageModel {

    public static String URL = "url";
    public static String Data = "data";
    /**
     * imageInputMethod
     */
    private String type = URL;
    /**
     * imageWidth
     */
    private int width;
    // image height
    private int height;
    // The map's address
    private String url;
    // image information
    private byte[] data;

    private int rowspan = 1;
    private int colspan = 1;


    public ImageModel() {

    }

    public ImageModel(byte[] data, int width, int height) {
        this.data = data;
        this.width = width;
        this.height = height;
        this.type = Data;
    }

    public ImageModel(String url, int width, int height) {
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

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

}
