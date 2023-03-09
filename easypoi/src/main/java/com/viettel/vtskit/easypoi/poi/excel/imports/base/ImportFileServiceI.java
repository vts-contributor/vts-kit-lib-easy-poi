package com.viettel.vtskit.easypoi.poi.excel.imports.base;

public interface ImportFileServiceI {

    /**
     * Upload File Returns the file address string
     * @param data
     * @return
     */
    String doUpload(byte[] data);

}
