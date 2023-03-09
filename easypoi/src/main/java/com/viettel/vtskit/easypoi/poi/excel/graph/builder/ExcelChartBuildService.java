/**
 *
 */
package com.viettel.vtskit.easypoi.poi.excel.graph.builder;

import com.viettel.vtskit.easypoi.poi.common.PoiCellUtil;
import com.viettel.vtskit.easypoi.poi.common.PoiExcelGraphDataUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import com.viettel.vtskit.easypoi.poi.excel.graph.constant.ExcelGraphElementType;
import com.viettel.vtskit.easypoi.poi.excel.graph.constant.ExcelGraphType;
import com.viettel.vtskit.easypoi.poi.excel.graph.entity.ExcelGraph;
import com.viettel.vtskit.easypoi.poi.excel.graph.entity.ExcelGraphElement;
import com.viettel.vtskit.easypoi.poi.excel.graph.entity.ExcelTitleCell;

import java.util.ArrayList;
import java.util.List;

/**
 * @Description
 */
public class ExcelChartBuildService {
    /**
     *
     * @param workbook
     * @param graphList
     * @param build Recalculate graph definitions with live data rows
     * @param append
     */
    public static void createExcelChart(Workbook workbook, List<ExcelGraph> graphList, Boolean build, Boolean append) {
        if (workbook != null && graphList != null) {
            //Set the default first sheet as a data item
            Sheet dataSouce = workbook.getSheetAt(0);
            if (dataSouce != null) {
                buildTitle(dataSouce, graphList);

                if (build) {
                    PoiExcelGraphDataUtil.buildGraphData(dataSouce, graphList);
                }
                if (append) {
                    buildExcelChart(dataSouce, dataSouce, graphList);
                } else {
                    Sheet sheet = workbook.createSheet("Graphical interface");
                    buildExcelChart(dataSouce, sheet, graphList);
                }
            }

        }
    }

    /**
     * Build the base graph
     * @param drawing
     * @param anchor
     * @param dataSourceSheet
     * @param graph
     */
    private static void buildExcelChart(Drawing drawing, ClientAnchor anchor, Sheet dataSourceSheet, ExcelGraph graph) {
        Chart chart = null;
        // TODO  The chart did not succeed
        //drawing.createChart(anchor);
        ChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        ExcelGraphElement categoryElement = graph.getCategory();

        ChartDataSource categoryChart;
        if (categoryElement != null && categoryElement.getElementType().equals(ExcelGraphElementType.STRING_TYPE)) {
            categoryChart = DataSources.fromStringCellRange(dataSourceSheet, new CellRangeAddress(categoryElement.getStartRowNum(), categoryElement.getEndRowNum(), categoryElement.getStartColNum(), categoryElement.getEndColNum()));
        } else {
            categoryChart = DataSources.fromNumericCellRange(dataSourceSheet, new CellRangeAddress(categoryElement.getStartRowNum(), categoryElement.getEndRowNum(), categoryElement.getStartColNum(), categoryElement.getEndColNum()));
        }

        List<ExcelGraphElement> valueList = graph.getValueList();
        List<ChartDataSource<Number>> chartValueList = new ArrayList<>();
        if (valueList != null && valueList.size() > 0) {
            for (ExcelGraphElement ele : valueList) {
                ChartDataSource<Number> source = DataSources.fromNumericCellRange(dataSourceSheet, new CellRangeAddress(ele.getStartRowNum(), ele.getEndRowNum(), ele.getStartColNum(), ele.getEndColNum()));
                chartValueList.add(source);
            }
        }

        if (graph.getGraphType().equals(ExcelGraphType.LINE_CHART)) {
            LineChartData data = chart.getChartDataFactory().createLineChartData();
            buildLineChartData(data, categoryChart, chartValueList, graph.getTitle());
            chart.plot(data, bottomAxis, leftAxis);
        } else {
            ScatterChartData data = chart.getChartDataFactory().createScatterChartData();
            buildScatterChartData(data, categoryChart, chartValueList, graph.getTitle());
            chart.plot(data, bottomAxis, leftAxis);
        }
    }


    /**
     * Build multiple graphical objects
     * @param dataSourceSheet
     * @param tragetSheet
     * @param graphList
     */
    private static void buildExcelChart(Sheet dataSourceSheet, Sheet tragetSheet, List<ExcelGraph> graphList) {
        int len = graphList.size();
        if (len == 1) {
            buildExcelChart(dataSourceSheet, tragetSheet, graphList.get(0));
        } else {
            int drawStart = 0;
            int drawEnd = 20;
            Drawing drawing = PoiExcelGraphDataUtil.getDrawingPatriarch(tragetSheet);
            for (int i = 0; i < len; i++) {
                ExcelGraph graph = graphList.get(i);
                ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, drawStart, 15, drawEnd);
                buildExcelChart(drawing, anchor, dataSourceSheet, graph);
                drawStart = drawStart + drawEnd;
                drawEnd = drawEnd + drawEnd;
            }
        }
    }


    /**
     * Build graphical objects
     * @param dataSourceSheet
     * @param tragetSheet
     * @param graph
     */
    private static void buildExcelChart(Sheet dataSourceSheet, Sheet tragetSheet, ExcelGraph graph) {
        Drawing drawing = PoiExcelGraphDataUtil.getDrawingPatriarch(tragetSheet);
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 15, 20);
        buildExcelChart(drawing, anchor, dataSourceSheet, graph);
    }


    /**
     * Build the Title
     * @param sheet
     * @param graph
     */
    private static void buildTitle(Sheet sheet, ExcelGraph graph) {
        int cellTitleLen = graph.getTitleCell().size();
        int titleLen = graph.getTitle().size();
        if (titleLen > 0) {

        } else {
            for (int i = 0; i < cellTitleLen; i++) {
                ExcelTitleCell titleCell = graph.getTitleCell().get(i);
                if (titleCell != null) {
                    graph.getTitle().add(PoiCellUtil.getCellValue(sheet, titleCell.getRow(), titleCell.getCol()));
                }
            }
        }
    }

    /**
     * Build the Title
     * @param sheet
     * @param graphList
     */
    private static void buildTitle(Sheet sheet, List<ExcelGraph> graphList) {
        if (graphList != null && graphList.size() > 0) {
            for (ExcelGraph graph : graphList) {
                if (graph != null) {
                    buildTitle(sheet, graph);
                }
            }
        }
    }

    /**
     *
     * @param data
     * @param categoryChart
     * @param chartValueList
     * @param title
     */
    private static void buildLineChartData(LineChartData data, ChartDataSource categoryChart, List<ChartDataSource<Number>> chartValueList, List<String> title) {
        if (chartValueList.size() == title.size()) {
            int len = title.size();
            for (int i = 0; i < len; i++) {
                //data.addSerie(categoryChart, chartValueList.get(i)).setTitle(title.get(i));
            }
        } else {
            int i = 0;
            for (ChartDataSource<Number> source : chartValueList) {
                String temp_title = title.get(i);
                if (StringUtils.isNotBlank(temp_title)) {
                    //data.addSerie(categoryChart, source).setTitle(_title);
                } else {
                    //data.addSerie(categoryChart, source);
                }
            }
        }
    }

    /**
     *
     * @param data
     * @param categoryChart
     * @param chartValueList
     * @param title
     */
    private static void buildScatterChartData(ScatterChartData data, ChartDataSource categoryChart, List<ChartDataSource<Number>> chartValueList, List<String> title) {
        if (chartValueList.size() == title.size()) {
            int len = title.size();
            for (int i = 0; i < len; i++) {
                data.addSerie(categoryChart, chartValueList.get(i)).setTitle(title.get(i));
            }
        } else {
            int i = 0;
            for (ChartDataSource<Number> source : chartValueList) {
                String temp_title = title.get(i);
                if (StringUtils.isNotBlank(temp_title)) {
                    data.addSerie(categoryChart, source).setTitle(temp_title);
                } else {
                    data.addSerie(categoryChart, source);
                }
            }
        }
    }


}
