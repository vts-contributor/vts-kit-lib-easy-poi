
package com.atviettelsolutions.easypoi.poi.excel.export.styler;

import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelForEachParams;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * The abstract interface provides two public methods
 *
 * @author caprocute
 */
public abstract class AbstractExcelExportStyler implements IExcelExportStyler {
    // Single
    protected CellStyle stringNoneStyle;
    protected CellStyle stringNoneWrapStyle;
    // INTERVAL ROWS
    protected CellStyle stringSeptailStyle;
    protected CellStyle stringSeptailWrapStyle;

    protected Workbook workbook;

    protected static final short STRING_FORMAT = (short) BuiltinFormats.getBuiltinFormat("TEXT");

    protected void createStyles(Workbook workbook) {
        this.stringNoneStyle = stringNoneStyle(workbook, false);
        this.stringNoneWrapStyle = stringNoneStyle(workbook, true);
        this.stringSeptailStyle = stringSeptailStyle(workbook, false);
        this.stringSeptailWrapStyle = stringSeptailStyle(workbook, true);
        this.workbook = workbook;
    }

    @Override
    public CellStyle getStyles(boolean noneStyler, ExcelExportEntity entity) {
        if (noneStyler && (entity == null || entity.isWrap())) {
            return stringNoneWrapStyle;
        }
        if (noneStyler) {
            return stringNoneStyle;
        }
        if (noneStyler == false && (entity == null || entity.isWrap())) {
            return stringSeptailWrapStyle;
        }
        return stringSeptailStyle;
    }

    public CellStyle stringNoneStyle(Workbook workbook, boolean isWarp) {
        return null;
    }

    public CellStyle stringSeptailStyle(Workbook workbook, boolean isWarp) {
        return null;
    }

    /**
     * Get template styles (used when cycling through columns)
     *
     * @param isSingle
     * @param excelForEachParams
     * @return
     */
    @Override
    public CellStyle getTemplateStyles(boolean isSingle, ExcelForEachParams excelForEachParams) {
        return null;
    }

}
