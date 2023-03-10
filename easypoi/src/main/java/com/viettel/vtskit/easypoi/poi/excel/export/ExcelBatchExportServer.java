package com.viettel.vtskit.easypoi.poi.excel.export;

import com.viettel.vtskit.easypoi.poi.common.PoiExcelGraphDataUtil;
import com.viettel.vtskit.easypoi.poi.common.PoiPublicUtil;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelTarget;
import com.viettel.vtskit.easypoi.poi.excel.entity.ExportParams;
import com.viettel.vtskit.easypoi.poi.excel.entity.enmus.ExcelType;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.viettel.vtskit.easypoi.poi.excel.entity.vo.PoiBaseConstants;
import com.viettel.vtskit.easypoi.poi.excel.export.styler.IExcelExportStyler;
import com.viettel.vtskit.easypoi.poi.exeption.ExcelExportException;
import com.viettel.vtskit.easypoi.poi.exeption.excel.enums.ExcelExportEnum;
import com.viettel.vtskit.easypoi.poi.handler.inter.IExcelExportServer;
import com.viettel.vtskit.easypoi.poi.handler.inter.IWriter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

import static com.viettel.vtskit.easypoi.poi.excel.ExcelExportUtil.USE_SXSSF_LIMIT;
import static com.viettel.vtskit.easypoi.poi.excel.ExcelImportCheckUtil.getAllExcelField;


/**
 * Batch insertion service is available
 */
public class ExcelBatchExportServer extends ExcelExportServer implements IWriter<Workbook> {

    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelBatchExportServer.class);

    private Workbook workbook;
    private Sheet sheet;
    private List<ExcelExportEntity> excelParams;
    private ExportParams entity;
    private int titleHeight;
    private Drawing patriarch;
    private short rowHeight;
    private int index;

    public void init(ExportParams entity, Class<?> pojoClass) {
        List<ExcelExportEntity> excelParams = createExcelExportEntityList(entity, pojoClass);
        init(entity, excelParams);
    }

    /**
     * Initialize the data
     *
     * @param entity      Export parameters
     * @param excelParams
     */
    public void init(ExportParams entity, List<ExcelExportEntity> excelParams) {
        LOGGER.debug("ExcelBatchExportServer only support SXSSFWorkbook");
        entity.setType(ExcelType.XSSF);
        workbook = new SXSSFWorkbook();
        this.entity = entity;
        this.excelParams = excelParams;
        super.type = entity.getType();
        createSheet(workbook, entity, excelParams);
        if (entity.getMaxNum() == 0) {
            entity.setMaxNum(USE_SXSSF_LIMIT);
        }
        insertDataToSheet(workbook, entity, excelParams, null, sheet);
    }

    public List<ExcelExportEntity> createExcelExportEntityList(ExportParams entity, Class<?> pojoClass) {
        try {
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            // Get all fields
            Field[] fileds = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            String targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass,
                    null);
            sortAllParams(excelParams);

            return excelParams;
        } catch (Exception e) {
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    public void createSheet(Workbook workbook, ExportParams entity, List<ExcelExportEntity> excelParams) {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel export start ,List<ExcelExportEntity> is {}", excelParams);
            LOGGER.debug("Excel version is {}",
                    entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || excelParams == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        try {
            try {
                sheet = workbook.createSheet(entity.getSheetName());
            } catch (Exception e) {
                // Repeated traversal, duplicate names occur, and a non-specified name Sheet is created
                sheet = workbook.createSheet();
            }
        } catch (Exception e) {
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    @Override
    protected void insertDataToSheet(Workbook workbook, ExportParams entity,
                                     List<ExcelExportEntity> entityList, Collection<? extends Map<?, ?>> dataSet,
                                     Sheet sheet) {
        try {
            dataHanlder = entity.getDataHanlder();
            if (dataHanlder != null && dataHanlder.getNeedHandlerFields() != null) {
                needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
            }
            // Create a table style
            setExcelExportStyler((IExcelExportStyler) entity.getStyle()
                    .getConstructor(Workbook.class).newInstance(workbook));
            patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            this.index = entity.isCreateHeadRows()
                    ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
            titleHeight = index;
            setCellWith(excelParams, sheet);
            setColumnHidden(excelParams, sheet);
            rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    public Workbook exportBigExcel(IExcelExportServer server, Object queryParams) {
        int page = 1;
        List<Object> list = server
                .selectListForExcelExport(queryParams, page++);
        while (list != null && list.size() > 0) {
            write(list);
            list = server.selectListForExcelExport(queryParams, page++);
        }
        return close();
    }

    @Override
    public Workbook get() {
        return this.workbook;
    }

    @Override
    public IWriter<Workbook> write(Collection data) {
        if (sheet.getLastRowNum() + data.size() > entity.getMaxNum()) {
            sheet = workbook.createSheet();
            index = 0;
        }
        Iterator<?> its = data.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            try {
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight, 0)[0];
            } catch (Exception e) {
                LOGGER.error(e.getMessage(), e);
                throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
            }
        }
        return this;
    }

    @Override
    public Workbook close() {
        if (entity.getFreezeCol() != 0) {
            sheet.createFreezePane(entity.getFreezeCol(), titleHeight, entity.getFreezeCol(), titleHeight);
        }
        mergeCells(sheet, excelParams, titleHeight);
        // Create totals
        addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);
        return workbook;
    }

    /**
     * Add an Index column
     */
    @Override
    public ExcelExportEntity indexExcelEntity(ExportParams entity) {
        ExcelExportEntity exportEntity = new ExcelExportEntity();
        //Guaranteed to be the first row
        exportEntity.setOrderNum(Integer.MIN_VALUE);
        exportEntity.setNeedMerge(true);
        exportEntity.setName(entity.getIndexName());
        exportEntity.setWidth(10);
        exportEntity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
        return exportEntity;
    }
}