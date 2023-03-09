package com.viettel.vtskit.easypoi.poi.handler.inter;

import java.util.List;

/**
 * Export data interface
 */
public interface IExcelExportServer {
    /**
     * Query data interface
     *
     * @param queryParams
     * @param page
     * @return
     */
    public List<Object> selectListForExcelExport(Object queryParams, int page);
}
