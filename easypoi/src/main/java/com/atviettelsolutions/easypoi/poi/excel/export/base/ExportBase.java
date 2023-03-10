
package com.atviettelsolutions.easypoi.poi.excel.export.base;

import com.atviettelsolutions.easypoi.core.util.ApplicationContextUtil;
import com.atviettelsolutions.easypoi.dict.service.AutoPoiDictServiceI;
import com.atviettelsolutions.easypoi.poi.excel.annotation.Excel;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelCollection;
import com.atviettelsolutions.easypoi.poi.excel.annotation.ExcelEntity;
import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelDataHandler;
import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelDictHandler;
import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import com.atviettelsolutions.easypoi.poi.excel.entity.ExportParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelExportEntity;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Export basic processing, do not design POI, only design objects, to ensure reusability
 *
 * @author caprocute
 */
public class ExportBase {

    protected IExcelDataHandler dataHanlder;

    protected IExcelDictHandler dictHandler;


    protected List<String> needHanlderList;

    /**
     * Create an export entity object
     *
     * @param field
     * @param targetId
     * @param pojoClass
     * @param getMethods
     * @return
     * @throws Exception
     */
    private ExcelExportEntity createExcelExportEntity(Field field, String targetId, Class<?> pojoClass, List<Method> getMethods) throws Exception {
        Excel excel = field.getAnnotation(Excel.class);
        ExcelExportEntity excelEntity = new ExcelExportEntity();
        excelEntity.setType(excel.type());
        getExcelField(targetId, field, excelEntity, excel, pojoClass);
        if (getMethods != null) {
            List<Method> newMethods = new ArrayList<Method>();
            newMethods.addAll(getMethods);
            newMethods.add(excelEntity.getMethod());
            excelEntity.setMethods(newMethods);
        }
        return excelEntity;
    }

    private Object formatValue(Object value, ExcelExportEntity entity) throws Exception {
        Date temp = null;
        if ("".equals(value)) {
            value = null;
        }
        if (value instanceof String && entity.getDatabaseFormat() != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getDatabaseFormat());
            temp = format.parse(value.toString());
        } else if (value instanceof Date) {
            temp = (Date) value;
        } else if (value instanceof LocalDateTime) {
            LocalDateTime ldt = (LocalDateTime) value;
            DateTimeFormatter format = DateTimeFormatter.ofPattern(entity.getFormat());
            return format.format(ldt);
        } else if (value instanceof LocalDate) {
            LocalDate ld = (LocalDate) value;
            DateTimeFormatter format = DateTimeFormatter.ofPattern(entity.getFormat());
            return format.format(ld);
        }
        if (temp != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            value = format.format(temp);
        }
        return value;
    }

    /**
     * Get all the fields that need to be exported
     *
     * @param exclusions
     * @param targetId
     * @param fields
     * @throws Exception
     */
    public void getAllExcelField(String[] exclusions, String targetId, Field[] fields,
                                 List<ExcelExportEntity> excelParams, Class<?> pojoClass, List<Method> getMethods) throws Exception {
        List<String> exclusionsList = exclusions != null ? Arrays.asList(exclusions) : null;
        ExcelExportEntity excelEntity;
        // Traverse the entire filed
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            // First determine whether it is a collection, in order to determine whether Java comes with an object, and then it is our own object
            if (PoiPublicUtil.isNotUserExcelUserThis(exclusionsList, field, targetId)) {
                continue;
            }
            // First of all, it is determined that Excel may customize the processing of special data users back to
            if (field.getAnnotation(Excel.class) != null) {
                excelParams.add(createExcelExportEntity(field, targetId, pojoClass, getMethods));
            } else if (PoiPublicUtil.isCollection(field.getType())) {
                ExcelCollection excel = field.getAnnotation(ExcelCollection.class);
                ParameterizedType pt = (ParameterizedType) field.getGenericType();
                Class<?> clz = (Class<?>) pt.getActualTypeArguments()[0];
                List<ExcelExportEntity> list = new ArrayList<ExcelExportEntity>();
                getAllExcelField(exclusions, StringUtils.isNotEmpty(excel.id()) ? excel.id() : targetId,
                        PoiPublicUtil.getClassFields(clz), list, clz, null);
                excelEntity = new ExcelExportEntity();
                excelEntity.setName(getExcelName(excel.name(), targetId));
                excelEntity.setOrderNum(getCellOrder(excel.orderNum(), targetId));
                excelEntity.setMethod(PoiPublicUtil.getMethod(field.getName(), pojoClass));
                excelEntity.setList(list);
                excelParams.add(excelEntity);
            } else {
                List<Method> newMethods = new ArrayList<Method>();
                if (getMethods != null) {
                    newMethods.addAll(getMethods);
                }
                newMethods.add(PoiPublicUtil.getMethod(field.getName(), pojoClass));
                ExcelEntity excel = field.getAnnotation(ExcelEntity.class);
                if (excel.show() == true) {
                    List<ExcelExportEntity> list = new ArrayList<ExcelExportEntity>();
                    // There is a design pit, the last parameter is null when exporting,
                    // that is, getgetMethods gets empty, and you need to set the level getmethod when importing
                    getAllExcelField(exclusions, StringUtils.isNotEmpty(excel.id()) ? excel.id() : targetId,
                            PoiPublicUtil.getClassFields(field.getType()), list, field.getType(), null);
                    excelEntity = new ExcelExportEntity();
                    excelEntity.setName(getExcelName(excel.name(), targetId));
                    excelEntity.setMethod(PoiPublicUtil.getMethod(field.getName(), pojoClass));
                    excelEntity.setList(list);
                    excelParams.add(excelEntity);
                } else {
                    getAllExcelField(exclusions, StringUtils.isNotEmpty(excel.id()) ? excel.id() :
                            targetId, PoiPublicUtil.getClassFields(field.getType()), excelParams, field.getType(), newMethods);
                }
            }
        }
    }

    /**
     * Get the order of this field
     *
     * @param orderNum
     * @param targetId
     * @return
     */
    public int getCellOrder(String orderNum, String targetId) {
        if (isInteger(orderNum) || targetId == null) {
            return Integer.valueOf(orderNum);
        }
        String[] arr = orderNum.split(",");
        String[] temp;
        for (String str : arr) {
            temp = str.split("_");
            if (targetId.equals(temp[1])) {
                return Integer.valueOf(temp[0]);
            }
        }
        return 0;
    }

    /**
     * Get the value of this cell, which provides some additional functions
     *
     * @param entity
     * @param obj
     * @return
     * @throws Exception
     */
    public Object getCellValue(ExcelExportEntity entity, Object obj) throws Exception {
        Object value;
        if (obj instanceof Map) {
            value = ((Map<?, ?>) obj).get(entity.getKey());
        } else {
            value = entity.getMethods() != null ? getFieldBySomeMethod(entity.getMethods(), obj) : entity.getMethod().invoke(obj, new Object[]{});
        }

        value = Optional.ofNullable(value).orElse("");
        if (StringUtils.isEmpty(value.toString())) {
            return "";
        }

        if (StringUtils.isNotEmpty(entity.getNumFormat()) && value != null) {
            value = new DecimalFormat(entity.getNumFormat()).format(value);
        }

        if (StringUtils.isNotEmpty(entity.getDict()) && dictHandler != null) {
            value = dictHandler.toName(entity.getDict(), obj, entity.getName(), value);
        }
        if (StringUtils.isNotEmpty(entity.getFormat())) {
            value = formatValue(value, entity);
        }
        if (entity.getReplace() != null && entity.getReplace().length > 0) {
            if (value == null) {
                value = "";
            }
            String oldVal = value.toString();
            if (entity.isMultiReplace()) {
                value = multiReplaceValue(entity.getReplace(), String.valueOf(value));
            } else {
                value = replaceValue(entity.getReplace(), String.valueOf(value));
            }

            if (oldVal.equals(value)) {

            }
        }
        if (needHanlderList != null && needHanlderList.contains(entity.getName())) {
            value = dataHanlder.exportHandler(obj, entity.getName(), value);
        }
        if (StringUtils.isNotEmpty(entity.getSuffix()) && value != null) {
            value = value + entity.getSuffix();
        }
        return value == null ? "" : value.toString();
    }

    /**
     * Gets the value of the collection
     *
     * @param entity
     * @param obj
     * @return
     * @throws Exception
     */
    public Collection<?> getListCellValue(ExcelExportEntity entity, Object obj) throws Exception {
        Object value;
        if (obj instanceof Map) {
            value = ((Map<?, ?>) obj).get(entity.getKey());
        } else {
            value = entity.getMethod().invoke(obj, new Object[]{});
            if (value instanceof Collection) {
                return (Collection<?>) value;
            } else {
                List list = new ArrayList();
                list.add(value);
                return list;
            }
        }
        return (Collection<?>) value;
    }

    /**
     * Annotation to the conversion of the exported object
     *
     * @param targetId
     * @param field
     * @param excelEntity
     * @param excel
     * @param pojoClass
     * @throws Exception
     */
    private void getExcelField(String targetId, Field field, ExcelExportEntity excelEntity, Excel excel, Class<?> pojoClass) throws Exception {
        excelEntity.setName(getExcelName(excel.name(), targetId));
        excelEntity.setWidth(excel.width());
        excelEntity.setHeight(excel.height());
        excelEntity.setNeedMerge(excel.needMerge());
        excelEntity.setMergeVertical(excel.mergeVertical());
        excelEntity.setMergeRely(excel.mergeRely());
        excelEntity.setReplace(excel.replace());
        excelEntity.setHyperlink(excel.isHyperlink());
        if (StringUtils.isNotEmpty(excel.dicCode())) {
            AutoPoiDictServiceI jeecgDictService = null;
            try {
                jeecgDictService = ApplicationContextUtil.getContext().getBean(AutoPoiDictServiceI.class);
            } catch (Exception e) {
            }
            if (jeecgDictService != null) {
                String[] dictReplace = jeecgDictService.queryDict(excel.dictTable(), excel.dicCode(), excel.dicText());
                if (excelEntity.getReplace() != null && dictReplace != null && dictReplace.length != 0) {
                    excelEntity.setReplace(dictReplace);
                }
            }
        }
        excelEntity.setOrderNum(getCellOrder(excel.orderNum(), targetId));
        excelEntity.setWrap(excel.isWrap());
        excelEntity.setExportImageType(excel.imageType());
        excelEntity.setSuffix(excel.suffix());
        excelEntity.setDatabaseFormat(excel.databaseFormat());
        excelEntity.setFormat(StringUtils.isNotEmpty(excel.exportFormat()) ? excel.exportFormat() : excel.format());
        excelEntity.setStatistics(excel.isStatistics());
        String fieldname = field.getName();
        excelEntity.setKey(fieldname);
        excelEntity.setNumFormat(excel.numFormat());
        excelEntity.setColumnHidden(excel.isColumnHidden());
        excelEntity.setMethod(PoiPublicUtil.getMethod(fieldname, pojoClass, excel.exportConvert()));
        excelEntity.setMultiReplace(excel.multiReplace());
        if (StringUtils.isNotEmpty(excel.groupName())) {
            excelEntity.setGroupName(excel.groupName());
            excelEntity.setColspan(true);
        }
    }

    /**
     * Determines the name displayed in this cell
     *
     * @param exportName
     * @param targetId
     * @return
     */
    public String getExcelName(String exportName, String targetId) {
        if (exportName.indexOf(",") < 0 || targetId == null) {
            return exportName;
        }
        String[] arr = exportName.split(",");
        for (String str : arr) {
            if (str.indexOf(targetId) != -1) {
                return str.split("_")[0];
            }
        }
        return null;
    }

    /**
     * Multiple reflections get values
     *
     * @param list
     * @param t
     * @return
     * @throws Exception
     */
    public Object getFieldBySomeMethod(List<Method> list, Object t) throws Exception {
        for (Method m : list) {
            if (t == null) {
                t = "";
                break;
            }
            t = m.invoke(t, new Object[]{});
        }
        return t;
    }

    /**
     * Get the row height based on the annotation
     *
     * @param excelParams
     * @return
     */
    public short getRowHeight(List<ExcelExportEntity> excelParams) {
        double maxHeight = 0;
        for (int i = 0; i < excelParams.size(); i++) {
            maxHeight = maxHeight > excelParams.get(i).getHeight() ? maxHeight : excelParams.get(i).getHeight();
            if (excelParams.get(i).getList() != null) {
                for (int j = 0; j < excelParams.get(i).getList().size(); j++) {
                    maxHeight = maxHeight > excelParams.get(i).getList().get(j).getHeight() ? maxHeight : excelParams.get(i).getList().get(j).getHeight();
                }
            }
        }
        return (short) (maxHeight * 50);
    }

    /**
     * Determines whether a string is an integer
     */
    public boolean isInteger(String value) {
        try {
            Integer.parseInt(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private Object replaceValue(String[] replace, String value) {
        String[] temp;
        for (String str : replace) {

            temp = getValueArr(str);

            if (value.equals(temp[1]) || value.replace("_", "---").equals(temp[1])) {
                value = temp[0];
                break;
            }
        }
        return value;
    }


    /**
     * If the value to be replaced is multi-option, there is a comma between each item, follow the following method
     */
    private Object multiReplaceValue(String[] replace, String value) {
        if (value.indexOf(",") > 0) {
            String[] radioVals = value.split(",");
            String[] temp;
            String result = "";
            for (int i = 0; i < radioVals.length; i++) {
                String radio = radioVals[i];
                for (String str : replace) {
                    temp = str.split("_");
                    temp = getValueArr(str);

                    if (radio.equals(temp[1]) || radio.replace("_", "---").equals(temp[1])) {
                        result = result.concat(temp[0]) + ",";
                        break;
                    }
                }
            }
            if (result.equals("")) {
                result = value;
            } else {
                result = result.substring(0, result.length() - 1);
            }
            return result;
        } else {
            return replaceValue(replace, value);
        }
    }

    /**
     * Sort fields based on user settings
     */
    public void sortAllParams(List<ExcelExportEntity> excelParams) {
        Collections.sort(excelParams);
        for (ExcelExportEntity entity : excelParams) {
            if (entity.getList() != null) {
                Collections.sort(entity.getList());
            }
        }
    }

    /**
     * Loop through the ExcelExportEntity collection Additional <br>configuration information
     * 1. Column sorting<br>
     * 2. Read the picture root path setting (if the field is the picture type and stored locally, set the disk path to get the full address export<br>).
     * 3. Multi-table header configuration (limited to a single table will take this logical processing)
     */
    public void reConfigExcelExportParams(List<ExcelExportEntity> excelParams, ExportParams exportParams) {
        Set<String> NameSet = new HashSet<String>();
        Map<String, List<String>> groupAndColumnList = new HashMap<String, List<String>>();
        Map<String, Integer> groupOrder = new HashMap<>();
        int index = -99;
        for (ExcelExportEntity entity : excelParams) {
            if (entity.getOrderNum() == 0) {
                entity.setOrderNum(index++);
            }
            if (entity.getExportImageType() == 3) {
                entity.setImageBasePath(exportParams.getImageBasePath());
            }
            if (entity.getList() != null) {
                Collections.sort(entity.getList());
            }
            String groupName = entity.getGroupName();
            if (StringUtils.isNotEmpty(groupName)) {
                List<String> ls = groupAndColumnList.get(groupName);
                if (ls == null) {
                    ls = new ArrayList<String>();
                    groupAndColumnList.put(groupName, ls);
                }
                ls.add(entity.getKey().toString());

                Integer order = groupOrder.get(groupName);
                if (order == null || entity.getOrderNum() < order) {
                    order = entity.getOrderNum();
                }
                groupOrder.put(groupName, order);
            }
        }

        for (String key : groupAndColumnList.keySet()) {
            ExcelExportEntity temp = new ExcelExportEntity(key);
            temp.setColspan(true);
            temp.setSubColumnList(groupAndColumnList.get(key));
            temp.setOrderNum(groupOrder.get(key));
            excelParams.add(temp);
        }
        Collections.sort(excelParams);
    }

    /**
     * The dictionary text contains multiple underscores, take the last one (to solve the null case)
     *
     * @param val
     * @return
     */
    public String[] getValueArr(String val) {
        int i = val.lastIndexOf("_");//The position of the last separator
        String[] c = new String[2];
        c[0] = val.substring(0, i); //label
        c[1] = val.substring(i + 1); //key
        return c;
    }
}
