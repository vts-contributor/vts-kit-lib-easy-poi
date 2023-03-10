package com.atviettelsolutions.easypoi.poi.excel.imports.sax.parse;

import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelReadRowHanlder;
import com.google.common.collect.Lists;
import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import com.atviettelsolutions.easypoi.poi.exeption.ExcelImportException;
import org.apache.commons.lang3.StringUtils;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelTarget;
import com.atviettelsolutions.easypoi.poi.excel.entity.ImportParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelCollectionParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelImportEntity;
import com.atviettelsolutions.easypoi.poi.excel.entity.sax.SaxReadCellEntity;
import com.atviettelsolutions.easypoi.poi.excel.imports.CellValueServer;
import com.atviettelsolutions.easypoi.poi.excel.imports.base.ImportBaseService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

/**
 * When the row reads the data
 *
 * @author caprocute
 */
@SuppressWarnings({"rawtypes", "unchecked"})
public class SaxRowRead extends ImportBaseService implements ISaxRowRead {

    private static final Logger LOGGER = LoggerFactory.getLogger(SaxRowRead.class);
    /**
     * The data that needs to be returned
     **/
    private List list;
    /**
     * The exported object
     **/
    private Class<?> pojoClass;
    /**
     * Import parameters
     **/
    private ImportParams params;
    /**
     * Table header correspondence
     **/
    private Map<Integer, String> titlemap = new HashMap<Integer, String>();
    /**
     * The current object
     **/
    private Object object = null;

    private Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();

    private List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();

    private String targetId;

    private CellValueServer cellValueServer;

    private IExcelReadRowHanlder hanlder;

    public SaxRowRead(Class<?> pojoClass, ImportParams params, IExcelReadRowHanlder hanlder) {
        list = Lists.newArrayList();
        this.params = params;
        this.pojoClass = pojoClass;
        cellValueServer = new CellValueServer();
        this.hanlder = hanlder;
        initParams(pojoClass, params);
    }

    private void initParams(Class<?> pojoClass, ImportParams params) {
        try {

            Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
            if (etarget != null) {
                targetId = etarget.value();
            }
            getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(e.getMessage());
        }

    }

    @Override
    public <T> List<T> getList() {
        return list;
    }

    @Override
    public void parse(int index, List<SaxReadCellEntity> datas) {
        try {
            if (datas == null || datas.size() == 0) {
                return;
            }
            // The header row is skipped
            if (index < params.getTitleRows()) {
                return;
            }
            // Header row
            if (index < params.getTitleRows() + params.getHeadRows()) {
                addHeadData(datas);
            } else {
                addListData(datas);
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(e.getMessage());
        }
    }

    /**
     * Collection element processing
     *
     * @param datas
     */
    private void addListData(List<SaxReadCellEntity> datas) throws Exception {
        // Determine whether it is a collection element or not, if so, continue to join the collection, or create a new object
        if ((datas.get(params.getKeyIndex()) == null || StringUtils.isEmpty(String.valueOf(datas.get(params.getKeyIndex()).getValue()))) && object != null) {
            for (ExcelCollectionParams param : excelCollection) {
                addListContinue(object, param, datas, titlemap, targetId, params);
            }
        } else {
            if (object != null && hanlder != null) {
                hanlder.hanlder(object);
            }
            object = PoiPublicUtil.createObject(pojoClass, targetId);
            SaxReadCellEntity entity;
            for (int i = 0, le = datas.size(); i < le; i++) {
                entity = datas.get(i);
                String titleString = (String) titlemap.get(i);
                if (excelParams.containsKey(titleString)) {
                    saveFieldValue(params, object, entity, excelParams, titleString);
                }
            }
            for (ExcelCollectionParams param : excelCollection) {
                addListContinue(object, param, datas, titlemap, targetId, params);
            }
            if (hanlder == null) {
                list.add(object);
            }
        }

    }

    /***
     * Continue adding elements to the List
     *
     * @param object
     * @param param
     * @param datas
     * @param titlemap
     * @param targetId
     * @param params
     */
    private void addListContinue(Object object, ExcelCollectionParams param, List<SaxReadCellEntity> datas, Map<Integer, String> titlemap, String targetId, ImportParams params) throws Exception {
        Collection collection = (Collection) PoiPublicUtil.getMethod(param.getName(), object.getClass()).invoke(object, new Object[]{});
        Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
        boolean isUsed = false;//
        for (int i = 0; i < datas.size(); i++) {
            String titleString = (String) titlemap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                saveFieldValue(params, entity, datas.get(i), param.getExcelParams(), titleString);
                isUsed = true;
            }
        }
        if (isUsed) {
            collection.add(entity);
        }
    }

    /**
     * Set the value
     *
     * @param params
     * @param object
     * @param entity
     * @param excelParams
     * @param titleString
     * @throws Exception
     */
    private void saveFieldValue(ImportParams params, Object object, SaxReadCellEntity entity, Map<String, ExcelImportEntity> excelParams, String titleString) throws Exception {
        Object value = cellValueServer.getValue(params.getDataHanlder(), object, entity, excelParams, titleString);
        setValues(excelParams.get(titleString), object, value);
    }

    /**
     * put
     *
     * @param datas
     */
    private void addHeadData(List<SaxReadCellEntity> datas) {
        for (int i = 0; i < datas.size(); i++) {
            if (StringUtils.isNotEmpty(String.valueOf(datas.get(i).getValue()))) {
                titlemap.put(i, String.valueOf(datas.get(i).getValue()));
            }
        }
    }
}
