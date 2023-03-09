
package com.viettel.vtskit.easypoi.poi.excel.export.styler;

import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelForEachParams;
import org.apache.poi.ss.usermodel.CellStyle;


/**
 * Excel export style interface
 *
 * @author caprocute
 */
public interface IExcelExportStyler {

    /**
     * List header style
     *
     * @param headerColor
     * @return
     */
    public CellStyle getHeaderStyle(short headerColor);

    /**
     * Heading style
     *
     * @param color
     * @return
     */
    public CellStyle getTitleStyle(short color);

    /**
     * Gets the style method
     *
     * @param noneStyler
     * @param entity
     * @return
     */
    public CellStyle getStyles(boolean noneStyler, ExcelExportEntity entity);

    /**
     * The style settings used by the template
     */
    public CellStyle getTemplateStyles(boolean isSingle, ExcelForEachParams excelForEachParams);

}
