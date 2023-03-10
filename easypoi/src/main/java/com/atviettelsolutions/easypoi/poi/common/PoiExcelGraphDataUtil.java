package com.atviettelsolutions.easypoi.poi.common;

import com.atviettelsolutions.easypoi.poi.excel.graph.entity.ExcelGraph;
import com.atviettelsolutions.easypoi.poi.excel.graph.entity.ExcelGraphElement;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;


import java.util.List;

/**
 * Build special data structures
 * @Description Big data export method [global]
 */
public class PoiExcelGraphDataUtil {
    /**
     * Build and get the last row of data and write it into the definition object
     * @param dataSourceSheet
     * @param graph
     */
    public static void buildGraphData(Sheet dataSourceSheet, ExcelGraph graph) {
        if (graph != null && graph.getCategory() != null && graph.getValueList() != null
                && graph.getValueList().size() > 0) {
            graph.getCategory().setEndRowNum(dataSourceSheet.getLastRowNum());
            for (ExcelGraphElement e : graph.getValueList()) {
                if (e != null) {
                    e.setEndRowNum(dataSourceSheet.getLastRowNum());
                }
            }
        }
    }

    /**
     * Build multiple graphics objects
     * @param dataSourceSheet
     * @param graphList
     */
    public static void buildGraphData(Sheet dataSourceSheet, List<ExcelGraph> graphList) {
        if (graphList != null && graphList.size() > 0) {
            for (ExcelGraph graph : graphList) {
                buildGraphData(dataSourceSheet, graph);
            }
        }
    }

    /**
     * Get the canvas, create one if not
     * @param sheet
     * @return
     */
    public static Drawing getDrawingPatriarch(Sheet sheet){
        if(sheet.getDrawingPatriarch() == null){
            sheet.createDrawingPatriarch();
        }
        return sheet.getDrawingPatriarch();
    }
}
