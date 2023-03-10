package com.atviettelsolutions.easypoi.poi.excel.export.base;

import com.atviettelsolutions.easypoi.poi.excel.export.styler.IExcelExportStyler;
import com.atviettelsolutions.easypoi.poi.common.MyX509TrustManager;
import com.atviettelsolutions.easypoi.poi.common.PoiMergeCellUtil;
import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import com.atviettelsolutions.easypoi.poi.excel.entity.enmus.ExcelType;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.atviettelsolutions.easypoi.poi.excel.entity.vo.PoiBaseConstants;
import com.atviettelsolutions.easypoi.poi.exeption.ExcelExportException;
import com.atviettelsolutions.easypoi.poi.exeption.excel.enums.ExcelExportEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.security.SecureRandom;
import java.text.DecimalFormat;
import java.util.*;

/**
 * Provide POI basic operation services
 *
 * @author caprocute
 */
public abstract class ExcelExportBase extends ExportBase {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelExportBase.class);

    private int currentIndex = 0;

    protected ExcelType type = ExcelType.HSSF;

    private Map<Integer, Double> statistics = new HashMap<Integer, Double>();

    private static final DecimalFormat DOUBLE_FORMAT = new DecimalFormat("######0.00");

    protected IExcelExportStyler excelExportStyler;


    /**
     * Create the main cells
     *
     * @param
     * @param rowHeight
     * @throws Exception
     */
    public int createCells(Drawing patriarch, int index, Object t,
                           List<ExcelExportEntity> excelParams, Sheet sheet, Workbook workbook, short rowHeight) throws Exception {
        ExcelExportEntity entity;
        Row row = sheet.createRow(index);
        row.setHeight(rowHeight);
        int maxHeight = 1, cellNum = 0;
        int indexKey = createIndexCell(row, index, excelParams.get(0));
        cellNum += indexKey;
        for (int k = indexKey, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.isSubColumn()) {
                continue;
            }
            if (entity.isMergeColumn()) {
                Map<String, Object> subColumnMap = new HashMap<>();
                List<String> mapKeys = entity.getSubColumnList();
                for (String subKey : mapKeys) {
                    Object subKeyValue = null;
                    if (t instanceof Map) {
                        subKeyValue = ((Map<?, ?>) t).get(subKey);
                    } else {
                        subKeyValue = PoiPublicUtil.getParamsValue(subKey, t);
                    }
                    subColumnMap.put(subKey, subKeyValue);
                }
                createListCells(patriarch, index, cellNum, subColumnMap, entity.getList(), sheet, workbook);
                cellNum += entity.getSubColumnList().size();
            } else if (entity.getList() != null) {
                Collection<?> list = getListCellValue(entity, t);
                int listC = 0;
                for (Object obj : list) {
                    createListCells(patriarch, index + listC, cellNum, obj, entity.getList(), sheet, workbook);
                    listC++;
                }
                cellNum += entity.getList().size();
                if (list != null && list.size() > maxHeight) {
                    maxHeight = list.size();
                }
            } else {
                Object value = getCellValue(entity, t);
                if (entity.getType() == 1) {
                    createStringCell(row, cellNum++, value == null ? "" : value.toString(), index % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity), entity);
                } else if (entity.getType() == 4) {
                    createNumericCell(row, cellNum++, value == null ? "" : value.toString(), index % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity), entity);
                } else {
                    createImageCell(patriarch, entity, row, cellNum++, value == null ? "" : value.toString(), t);
                }

                if (entity.isHyperlink()) {
                    row.getCell(cellNum - 1).setHyperlink(dataHanlder.getHyperlink(row.getSheet().getWorkbook().getCreationHelper(), t, entity.getName(), value));
                }
            }
        }
        // Merge the cells that need to be merged
        cellNum = 0;
        for (int k = indexKey, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            if (entity.getList() != null) {
                cellNum += entity.getList().size();
            } else if (entity.isNeedMerge()) {
                for (int i = index + 1; i < index + maxHeight; i++) {
                    sheet.getRow(i).createCell(cellNum);
                    sheet.getRow(i).getCell(cellNum).setCellStyle(getStyles(false, entity));
                }
                try {
                    sheet.addMergedRegion(new CellRangeAddress(index, index + maxHeight - 1, cellNum, cellNum));
                } catch (IllegalArgumentException e) {
                    LOGGER.error("Merge cell error logs：" + e.getMessage());
                    e.fillInStackTrace();
                }
                cellNum++;
            }
        }
        return maxHeight;

    }

    /**
     * Obtain image data through HTTPS address
     *
     * @param imagePath
     * @return
     * @throws Exception
     */
    private byte[] getImageDataByHttps(String imagePath) throws Exception {
        SSLContext sslcontext = SSLContext.getInstance("SSL", "SunJSSE");
        sslcontext.init(null, new TrustManager[]{new MyX509TrustManager()}, new SecureRandom());
        URL url = new URL(imagePath);
        HttpsURLConnection conn = (HttpsURLConnection) url.openConnection();
        conn.setSSLSocketFactory(sslcontext.getSocketFactory());
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5 * 1000);
        InputStream inStream = conn.getInputStream();
        byte[] value = readInputStream(inStream);
        return value;
    }

    /**
     * Obtain image data through HTTP address
     *
     * @param imagePath
     * @return
     * @throws Exception
     */
    private byte[] getImageDataByHttp(String imagePath) throws Exception {
        URL url = new URL(imagePath);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5 * 1000);
        InputStream inStream = conn.getInputStream();
        byte[] value = readInputStream(inStream);
        return value;
    }

    /**
     * Image type Cell
     *
     * @param patriarch
     * @param entity
     * @param row
     * @param i
     * @param imagePath
     * @param obj
     * @throws Exception
     */
    public void createImageCell(Drawing patriarch, ExcelExportEntity entity, Row row, int i, String imagePath, Object obj) throws Exception {
        row.setHeight((short) (50 * entity.getHeight()));
        row.createCell(i);
        ClientAnchor anchor;
        if (type.equals(ExcelType.HSSF)) {
            anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) i, row.getRowNum(), (short) (i + 1), row.getRowNum() + 1);
        } else {
            anchor = new XSSFClientAnchor(0, 0, 0, 0, (short) i, row.getRowNum(), (short) (i + 1), row.getRowNum() + 1);
        }

        if (StringUtils.isEmpty(imagePath)) {
            return;
        }

        int imageType = entity.getExportImageType();
        byte[] value = null;
        if (imageType == 2) {
            //Original logic 2
            value = (byte[]) (entity.getMethods() != null ? getFieldBySomeMethod(entity.getMethods(), obj) : entity.getMethod().invoke(obj, new Object[]{}));
        } else if (imageType == 4 || imagePath.startsWith("http")) {
            //Added logical network images 4
            try {
                if (imagePath.indexOf(",") != -1) {
                    if (imagePath.startsWith(",")) {
                        imagePath = imagePath.substring(1);
                    }
                    String[] images = imagePath.split(",");
                    imagePath = images[0];
                }
                if (imagePath.startsWith("https")) {
                    value = getImageDataByHttps(imagePath);
                } else {
                    value = getImageDataByHttp(imagePath);
                }
            } catch (Exception exception) {
                LOGGER.warn(exception.getMessage());
                //exception.printStackTrace();
            }
        } else {
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            BufferedImage bufferImg;
            String path = null;
            if (imageType == 1) {
                //Original logic 1
                path = PoiPublicUtil.getWebRootPath(imagePath);
                LOGGER.debug("--- createImageCell getWebRootPath ----filePath--- " + path);
                path = path.replace("WEB-INF/classes/", "");
                path = path.replace("file:/", "");
            } else if (imageType == 3) {
                //Added Logic Local Image 3
                if (StringUtils.isNotBlank(entity.getImageBasePath())) {
                    if (!entity.getImageBasePath().endsWith(File.separator) && !imagePath.startsWith(File.separator)) {
                        path = entity.getImageBasePath() + File.separator + imagePath;
                    } else {
                        path = entity.getImageBasePath() + imagePath;
                    }
                } else {
                    path = imagePath;
                }
            }
            try {
                bufferImg = ImageIO.read(new File(path));
                ImageIO.write(bufferImg, imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length()), byteArrayOut);
                value = byteArrayOut.toByteArray();
            } catch (Exception e) {
                LOGGER.error(e.getMessage());
            }
        }
        if (value != null) {
            patriarch.createPicture(anchor, row.getSheet().getWorkbook().addPicture(value, getImageType(value)));
        }


    }

    /**
     * in Stream reads into a byte array
     *
     * @param inStream
     * @return
     * @throws Exception
     */
    private byte[] readInputStream(InputStream inStream) throws Exception {
        if (inStream == null) {
            return null;
        }
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len = 0;
        //The length of the string for each read, if -1, means that all are read
        while ((len = inStream.read(buffer)) != -1) {
            outStream.write(buffer, 0, len);
        }
        inStream.close();
        return outStream.toByteArray();
    }

    private int createIndexCell(Row row, int index, ExcelExportEntity excelExportEntity) {
        // hieuhk1 - equal risk here :))
        if (excelExportEntity.getName().equals("serial number") && PoiBaseConstants.IS_ADD_INDEX.equals(excelExportEntity.getFormat())) {
            createStringCell(row, 0, currentIndex + "", index % 2 == 0 ? getStyles(false, null) : getStyles(true, null), null);
            currentIndex = currentIndex + 1;
            return 1;
        }
        return 0;
    }

    /**
     * Create the individual cells after the List
     *
     * @param patriarch
     * @param index
     * @param cellNum
     * @param obj
     * @param excelParams
     * @param sheet
     * @param workbook
     * @throws Exception
     */
    public void createListCells(Drawing patriarch, int index, int cellNum, Object obj, List<ExcelExportEntity> excelParams, Sheet sheet, Workbook workbook) throws Exception {
        ExcelExportEntity entity;
        Row row;
        if (sheet.getRow(index) == null) {
            row = sheet.createRow(index);
            row.setHeight(getRowHeight(excelParams));
        } else {
            row = sheet.getRow(index);
        }
        for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
            entity = excelParams.get(k);
            Object value = getCellValue(entity, obj);
            if (entity.getType() == 1) {
                createStringCell(row, cellNum++, value == null ? "" : value.toString(), row.getRowNum() % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity), entity);
                if (entity.isHyperlink()) {
                    row.getCell(cellNum - 1).setHyperlink(dataHanlder.getHyperlink(row.getSheet().getWorkbook().getCreationHelper(), obj, entity.getName(), value));
                }
            } else if (entity.getType() == 4) {
                createNumericCell(row, cellNum++, value == null ? "" : value.toString(), index % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity), entity);
                if (entity.isHyperlink()) {
                    row.getCell(cellNum - 1).setHyperlink(dataHanlder.getHyperlink(row.getSheet().getWorkbook().getCreationHelper(), obj, entity.getName(), value));
                }
            } else {
                createImageCell(patriarch, entity, row, cellNum++, value == null ? "" : value.toString(), obj);
            }
        }
    }

    public void createNumericCell(Row row, int index, String text, CellStyle style, ExcelExportEntity entity) {
        Cell cell = row.createCell(index);
        if (StringUtils.isEmpty(text)) {
            cell.setCellValue("");
            cell.setCellType(CellType.BLANK);
        } else {
            cell.setCellValue(Double.parseDouble(text));
            cell.setCellType(CellType.NUMERIC);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
        addStatisticsData(index, text, entity);
    }

    /**
     * Create a text type Cell
     *
     * @param row
     * @param index
     * @param text
     * @param style
     * @param entity
     */
    public void createStringCell(Row row, int index, String text, CellStyle style, ExcelExportEntity entity) {
        Cell cell = row.createCell(index);
        if (style != null && style.getDataFormat() > 0 && style.getDataFormat() < 12) {
            cell.setCellValue(Double.parseDouble(text));
            cell.setCellType(CellType.NUMERIC);
        } else {
            RichTextString Rtext;
            if (type.equals(ExcelType.HSSF)) {
                Rtext = new HSSFRichTextString(text);
            } else {
                Rtext = new XSSFRichTextString(text);
            }
            cell.setCellValue(Rtext);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
        addStatisticsData(index, text, entity);
    }

    /**
     * Create statistical lines
     *
     * @param styles
     * @param sheet
     */
    public void addStatisticsRow(CellStyle styles, Sheet sheet) {
        if (statistics.size() > 0) {
            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            Set<Integer> keys = statistics.keySet();
            createStringCell(row, 0, "total", styles, null);
            for (Integer key : keys) {
                createStringCell(row, key, DOUBLE_FORMAT.format(statistics.get(key)), styles, null);
            }
            statistics.clear();
        }

    }

    /**
     * totalStatistics
     *
     * @param index
     * @param text
     * @param entity
     */
    private void addStatisticsData(Integer index, String text, ExcelExportEntity entity) {
        if (entity != null && entity.isStatistics()) {
            Double temp = 0D;
            if (!statistics.containsKey(index)) {
                statistics.put(index, temp);
            }
            try {
                temp = Double.valueOf(text);
            } catch (NumberFormatException e) {
            }
            statistics.put(index, statistics.get(index) + temp);
        }
    }

    /**
     * Gets the total length of the fields for the exported report
     *
     * @param excelParams
     * @return
     */
    public int getFieldWidth(List<ExcelExportEntity> excelParams) {
        int length = -1;// Cells are calculated from 0
        for (ExcelExportEntity entity : excelParams) {
            if (entity.getGroupName() != null) {
                continue;
            } else if (entity.getSubColumnList() != null && entity.getSubColumnList().size() > 0) {
                length += entity.getSubColumnList().size();
            } else {
                length += entity.getList() != null ? entity.getList().size() : 1;
            }
        }
        return length;
    }

    /**
     * Get the image type and set the image insertion type
     *
     * @param value
     * @return
     * @author caprocute
     */
    public int getImageType(byte[] value) {
        String type = PoiPublicUtil.getFileExtendName(value);
        if (type.equalsIgnoreCase("JPG")) {
            return Workbook.PICTURE_TYPE_JPEG;
        } else if (type.equalsIgnoreCase("PNG")) {
            return Workbook.PICTURE_TYPE_PNG;
        }
        return Workbook.PICTURE_TYPE_JPEG;
    }

    private Map<Integer, int[]> getMergeDataMap(List<ExcelExportEntity> excelParams) {
        Map<Integer, int[]> mergeMap = new HashMap<Integer, int[]>();
        // Set the parameter order in preparation for merging cells later
        int i = 0;
        for (ExcelExportEntity entity : excelParams) {
            if (entity.isMergeVertical()) {
                mergeMap.put(i, entity.getMergeRely());
            }
            if (entity.getList() != null) {
                for (ExcelExportEntity inner : entity.getList()) {
                    if (inner.isMergeVertical()) {
                        mergeMap.put(i, inner.getMergeRely());
                    }
                    i++;
                }
            } else {
                i++;
            }
        }
        return mergeMap;
    }

    /**
     * Gets the style
     *
     * @param entity
     * @param needOne
     * @return
     */
    public CellStyle getStyles(boolean needOne, ExcelExportEntity entity) {
        return excelExportStyler.getStyles(needOne, entity);
    }

    /**
     * Merge cells
     *
     * @param sheet
     * @param excelParams
     * @param titleHeight
     */
    public void mergeCells(Sheet sheet, List<ExcelExportEntity> excelParams, int titleHeight) {
        Map<Integer, int[]> mergeMap = getMergeDataMap(excelParams);
        PoiMergeCellUtil.mergeCells(sheet, mergeMap, titleHeight);
    }

    public void setCellWith(List<ExcelExportEntity> excelParams, Sheet sheet) {
        int index = 0;
        for (int i = 0; i < excelParams.size(); i++) {
            if (excelParams.get(i).getList() != null) {
                List<ExcelExportEntity> list = excelParams.get(i).getList();
                for (int j = 0; j < list.size(); j++) {
                    sheet.setColumnWidth(index, (int) (256 * list.get(j).getWidth()));
                    index++;
                }
            } else {
                sheet.setColumnWidth(index, (int) (256 * excelParams.get(i).getWidth()));
                index++;
            }
        }
    }

    /**
     * setColumnHidden
     *
     * @param excelParams
     * @param sheet
     */
    public void setColumnHidden(List<ExcelExportEntity> excelParams, Sheet sheet) {
        int index = 0;
        for (int i = 0; i < excelParams.size(); i++) {
            if (excelParams.get(i).getList() != null) {
                List<ExcelExportEntity> list = excelParams.get(i).getList();
                for (int j = 0; j < list.size(); j++) {
                    sheet.setColumnHidden(index, list.get(j).isColumnHidden());
                    index++;
                }
            } else {
                sheet.setColumnHidden(index, excelParams.get(i).isColumnHidden());
                index++;
            }
        }
    }

    public void setCurrentIndex(int currentIndex) {
        this.currentIndex = currentIndex;
    }

    public void setExcelExportStyler(IExcelExportStyler excelExportStyler) {
        this.excelExportStyler = excelExportStyler;
    }

    public IExcelExportStyler getExcelExportStyler() {
        return excelExportStyler;
    }

    /**
     * Create a cell, returning the maximum height and number of cells
     *
     * @param patriarch
     * @param index
     * @param t
     * @param excelParams
     * @param sheet
     * @param workbook
     * @param rowHeight
     * @param cellNum
     * @return
     */
    public int[] createCells(Drawing patriarch, int index, Object t, List<ExcelExportEntity> excelParams, Sheet sheet, Workbook workbook, short rowHeight, int cellNum) {
        try {
            ExcelExportEntity entity;
            Row row = sheet.getRow(index) == null ? sheet.createRow(index) : sheet.getRow(index);
            if (rowHeight != -1) {
                row.setHeight(rowHeight);
            }
            int maxHeight = 1, listMaxHeight = 1;
            // Merge the cells that need to be merged
            int margeCellNum = cellNum;
            int indexKey = 0;
            if (excelParams != null && !excelParams.isEmpty()) {
                indexKey = createIndexCell(row, index, excelParams.get(0));
            }
            cellNum += indexKey;
            for (int k = indexKey, paramSize = excelParams.size(); k < paramSize; k++) {
                entity = excelParams.get(k);
                //Regardless of whether the data is empty or not, the data for that column should be jumped over
                if (entity.getList() != null) {
                    Collection<?> list = getListCellValue(entity, t);
                    int tmpListHeight = 0;
                    if (list != null && list.size() > 0) {
                        int tempCellNum = 0;
                        for (Object obj : list) {
                            int[] temp = createCells(patriarch, index + tmpListHeight, obj, entity.getList(), sheet, workbook, rowHeight, cellNum);
                            tempCellNum = temp[1];
                            tmpListHeight += temp[0];
                        }
                        cellNum = tempCellNum;
                        listMaxHeight = Math.max(listMaxHeight, tmpListHeight);
                    } else {
                        cellNum = cellNum + getListCellSize(entity.getList());
                    }
                } else {
                    Object value = getCellValue(entity, t);
                    if (entity.getType() == 1) {
                        createStringCell(row, cellNum++, value == null ? "" : value.toString(), index % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity), entity);

                    } else if (entity.getType() == 4) {
                        createNumericCell(row, cellNum++, value == null ? "" : value.toString(), index % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity), entity);
                    } else {
                        createImageCell(patriarch, entity, row, cellNum++, value == null ? "" : value.toString(), t);
                    }
                    if (entity.isHyperlink()) {
                        row.getCell(cellNum - 1).setHyperlink(dataHanlder.getHyperlink(row.getSheet().getWorkbook().getCreationHelper(), t, entity.getName(), value));
                    }
                }
            }
            maxHeight += listMaxHeight - 1;
            if (indexKey == 1 && excelParams.get(1).isNeedMerge()) {
                excelParams.get(0).setNeedMerge(true);
            }
            for (int k = indexKey, paramSize = excelParams.size(); k < paramSize; k++) {
                entity = excelParams.get(k);
                if (entity.getList() != null) {
                    margeCellNum += entity.getList().size();
                } else if (entity.isNeedMerge() && maxHeight > 1) {
                    for (int i = index + 1; i < index + maxHeight; i++) {
                        if (sheet.getRow(i) == null) {
                            sheet.createRow(i);
                        }
                        sheet.getRow(i).createCell(margeCellNum);
                        sheet.getRow(i).getCell(margeCellNum).setCellStyle(getStyles(false, entity));
                    }
                    PoiMergeCellUtil.addMergedRegion(sheet, index, index + maxHeight - 1, margeCellNum, margeCellNum);
                    margeCellNum++;
                }
            }
            return new int[]{maxHeight, cellNum};
        } catch (Exception e) {
            LOGGER.error("excel cell export error ,data is :{}", ReflectionToStringBuilder.toString(t));
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    /**
     * Gets the width of the collection
     *
     * @param list
     * @return
     */
    protected int getListCellSize(List<ExcelExportEntity> list) {
        int cellSize = 0;
        for (ExcelExportEntity ee : list) {
            if (ee.getList() != null) {
                cellSize += getListCellSize(ee.getList());
            } else {
                cellSize++;
            }
        }
        return cellSize;
    }
}
