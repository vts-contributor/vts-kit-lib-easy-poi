package com.viettel.vtskit.easypoi.poi.common;

import com.viettel.vtskit.easypoi.poi.excel.annotation.Excel;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelCollection;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelEntity;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelIgnore;
import com.viettel.vtskit.easypoi.poi.excel.entity.vo.PoiBaseConstants;
import com.viettel.vtskit.easypoi.poi.word.entity.WordImageEntity;
import com.viettel.vtskit.easypoi.poi.word.entity.params.ExcelListEntity;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ClassUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URISyntaxException;
import java.util.*;

/**
 * AutoPoi public base class for
 */
public final class PoiPublicUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(PoiPublicUtil.class);

    private PoiPublicUtil() {

    }

    @SuppressWarnings({"unchecked"})
    public static <K, V> Map<K, V> mapFor(Object... mapping) {
        Map<K, V> map = new HashMap<K, V>();
        for (int i = 0; i < mapping.length; i += 2) {
            map.put((K) mapping[i], (V) mapping[i + 1]);
        }
        return map;
    }

    /**
     * completely create an object
     *
     * @param clazz
     * @return
     */
    public static Object createObject(Class<?> clazz, String targetId) {
        Object obj = null;
        Method setMethod;
        try {
            if (clazz.equals(Map.class)) {
                return new HashMap<String, Object>();
            }
            obj = clazz.newInstance();
            Field[] fields = getClassFields(clazz);
            for (Field field : fields) {
                if (isNotUserExcelUserThis(null, field, targetId)) {
                    continue;
                }
                if (isCollection(field.getType())) {
                    ExcelCollection collection = field.getAnnotation(ExcelCollection.class);
                    setMethod = getMethod(field.getName(), clazz, field.getType());
                    setMethod.invoke(obj, collection.type().newInstance());
                } else if (!isJavaClass(field)) {
                    setMethod = getMethod(field.getName(), clazz, field.getType());
                    setMethod.invoke(obj, createObject(field.getType(), targetId));
                }
            }

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new RuntimeException("create object exception");
        }
        return obj;

    }

    /**
     * Get the class, including the parent class
     *
     * @param clazz
     * @return
     */
    public static Field[] getClassFields(Class<?> clazz) {
        List<Field> list = new ArrayList<Field>();
        Field[] fields;
        do {
            fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                list.add(fields[i]);
            }
            clazz = clazz.getSuperclass();
        } while (clazz != Object.class && clazz != null);
        return list.toArray(fields);
    }

    /**
     * @param photoByte
     * @return
     */
    public static String getFileExtendName(byte[] photoByte) {
        String strFileExtendName = "JPG";
        if ((photoByte[0] == 71) && (photoByte[1] == 73) && (photoByte[2] == 70) && (photoByte[3] == 56) && ((photoByte[4] == 55) || (photoByte[4] == 57)) && (photoByte[5] == 97)) {
            strFileExtendName = "GIF";
        } else if ((photoByte[6] == 74) && (photoByte[7] == 70) && (photoByte[8] == 73) && (photoByte[9] == 70)) {
            strFileExtendName = "JPG";
        } else if ((photoByte[0] == 66) && (photoByte[1] == 77)) {
            strFileExtendName = "BMP";
        } else if ((photoByte[1] == 80) && (photoByte[2] == 78) && (photoByte[3] == 71)) {
            strFileExtendName = "PNG";
        }
        return strFileExtendName;
    }

    /**
     * Get the GET method
     *
     * @param name
     * @param pojoClass
     * @return
     * @throws Exception
     */
    public static Method getMethod(String name, Class<?> pojoClass) throws Exception {
        StringBuffer getMethodName = new StringBuffer(PoiBaseConstants.GET);
        getMethodName.append(name.substring(0, 1).toUpperCase());
        getMethodName.append(name.substring(1));
        Method method = null;
        try {
            method = pojoClass.getMethod(getMethodName.toString(), new Class[]{});
        } catch (Exception e) {
            method = pojoClass.getMethod(getMethodName.toString().replace(PoiBaseConstants.GET, PoiBaseConstants.IS), new Class[]{});
        }
        return method;
    }

    /**
     * Get the SET method
     *
     * @param name
     * @param pojoClass
     * @param type
     * @return
     * @throws Exception
     */
    public static Method getMethod(String name, Class<?> pojoClass, Class<?> type) throws Exception {
        StringBuffer getMethodName = new StringBuffer(PoiBaseConstants.SET);
        getMethodName.append(name.substring(0, 1).toUpperCase());
        getMethodName.append(name.substring(1));
        return pojoClass.getMethod(getMethodName.toString(), new Class[]{type});
    }

    /**
     * Obtain the get method Use the EXCEL annotation exportConvert to determine whether to support value conversion
     *
     * @param name
     * @param pojoClass
     * @param convert
     * @return
     * @throws Exception
     */
    public static Method getMethod(String name, Class<?> pojoClass, boolean convert) throws Exception {
        StringBuffer getMethodName = new StringBuffer();
        if (convert) {
            getMethodName.append(PoiBaseConstants.CONVERT);
        }
        getMethodName.append(PoiBaseConstants.GET);
        getMethodName.append(name.substring(0, 1).toUpperCase());
        getMethodName.append(name.substring(1));
        Method method = null;
        try {
            method = pojoClass.getMethod(getMethodName.toString(), new Class[]{});
        } catch (Exception e) {
            method = pojoClass.getMethod(getMethodName.toString().replace(PoiBaseConstants.GET, PoiBaseConstants.IS), new Class[]{});
        }
        return method;
    }

    /**
     * Get the set method Use the EXCEL annotation importConvert to determine whether to support value conversion
     *
     * @param name
     * @param pojoClass
     * @param type
     * @param convert
     * @return
     * @throws Exception
     */
    public static Method getMethod(String name, Class<?> pojoClass, Class<?> type, boolean convert) throws Exception {
        StringBuffer setMethodName = new StringBuffer();
        if (convert) {
            setMethodName.append(PoiBaseConstants.CONVERT);
        }
        setMethodName.append(PoiBaseConstants.SET);
        setMethodName.append(name.substring(0, 1).toUpperCase());
        setMethodName.append(name.substring(1));
        return pojoClass.getMethod(setMethodName.toString(), new Class[]{type});
    }

    /**
     * Get Excel 2003 Picture
     *
     * @param sheet
     * @param workbook
     * @return Map key: picture cell index (1_1) String, value: picture stream PictureData
     */
    public static Map<String, PictureData> getSheetPictrues03(HSSFSheet sheet, HSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (!pictures.isEmpty()) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = String.valueOf(anchor.getRow1()) + "_" + String.valueOf(anchor.getCol1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

    /**
     * Get Excel 2007 Image
     *
     * @param sheet
     * @param workbook
     * @return Map key: picture cell index (1_1) String, value: picture stream PictureData
     */
    public static Map<String, PictureData> getSheetPictrues07(XSSFSheet sheet, XSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = ctMarker.getRow() + "_" + ctMarker.getCol();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }
        return sheetIndexPicMap;
    }

    public static String getWebRootPath(String filePath) {
        try {
            String path = null;
            try {
                path = PoiPublicUtil.class.getClassLoader().getResource("").toURI().getPath();
            } catch (URISyntaxException e) {
                //e.printStackTrace();
            } catch (NullPointerException e) {
                path = PoiPublicUtil.class.getProtectionDomain().getCodeSource().getLocation().getPath();
            }
            if (path == null || path == "") {
                //Solve the problem of springboot startup mode, the upload path is empty
                path = ClassUtils.getDefaultClassLoader().getResource("").getPath();
            }
            //update-end--Author:zhangdaihao  Date:20190424 for：解决springboot 启动模式，上传路径获取为空问题----------------------
            LOGGER.debug("--- getWebRootPath ----filePath--- " + path);
            path = path.replace("WEB-INF/classes/", "");
            path = path.replace("file:/", "");
            LOGGER.debug("--- path---  " + path);
            LOGGER.debug("--- filePath---  " + filePath);
            return path + filePath;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Judging whether it is an implementation class of a collection
     *
     * @param clazz
     * @return
     */
    public static boolean isCollection(Class<?> clazz) {
        return Collection.class.isAssignableFrom(clazz);
    }

    /**
     * Is it java basic class
     *
     * @param field
     * @return
     */
    public static boolean isJavaClass(Field field) {
        Class<?> fieldType = field.getType();
        boolean isBaseClass = false;
        if (fieldType.isArray()) {
            isBaseClass = false;
        } else if (fieldType.isPrimitive() || fieldType.getPackage() == null || fieldType.getPackage().getName().equals("java.lang") || fieldType.getPackage().getName().equals("java.math") || fieldType.getPackage().getName().equals("java.sql") || fieldType.getPackage().getName().equals("java.util")) {
            isBaseClass = true;
        }
        return isBaseClass;
    }

    /**
     * Judging whether not to be in this excel operation
     *
     * @param
     * @param field
     * @param targetId
     * @return
     */
    public static boolean isNotUserExcelUserThis(List<String> exclusionsList, Field field, String targetId) {
        boolean boo = true;
        if (field.getAnnotation(ExcelIgnore.class) != null) {
            boo = true;
        } else if (boo && field.getAnnotation(ExcelCollection.class) != null && isUseInThis(field.getAnnotation(ExcelCollection.class).name(), targetId) && (exclusionsList == null || !exclusionsList.contains(field.getAnnotation(ExcelCollection.class).name()))) {
            boo = false;
        } else if (boo && field.getAnnotation(Excel.class) != null && isUseInThis(field.getAnnotation(Excel.class).name(), targetId) && (exclusionsList == null || !exclusionsList.contains(field.getAnnotation(Excel.class).name()))) {
            boo = false;
        } else if (boo && field.getAnnotation(ExcelEntity.class) != null && isUseInThis(field.getAnnotation(ExcelEntity.class).name(), targetId) && (exclusionsList == null || !exclusionsList.contains(field.getAnnotation(ExcelEntity.class).name()))) {
            boo = false;
        }
        return boo;
    }

    /**
     * Determine whether to use
     *
     * @param exportName
     * @param targetId
     * @return
     */
    private static boolean isUseInThis(String exportName, String targetId) {
        return targetId == null || exportName.equals("") || exportName.indexOf("_") < 0 || exportName.indexOf(targetId) != -1;
    }

    private static Integer getImageType(String type) {
        if (type.equalsIgnoreCase("JPG") || type.equalsIgnoreCase("JPEG")) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        }
        if (type.equalsIgnoreCase("GIF")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        }
        if (type.equalsIgnoreCase("BMP")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        }
        if (type.equalsIgnoreCase("PNG")) {
            return XWPFDocument.PICTURE_TYPE_PNG;
        }
        return XWPFDocument.PICTURE_TYPE_JPEG;
    }

    /**
     * Return stream and image types
     *
     * @param entity
     * @return (byte[]) isAndType[0],(Integer)isAndType[1]
     * @throws Exception
     */
    public static Object[] getIsAndType(WordImageEntity entity) throws Exception {
        Object[] result = new Object[2];
        String type;
        if (entity.getType().equals(WordImageEntity.URL)) {
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            BufferedImage bufferImg;
            String path = Thread.currentThread().getContextClassLoader().getResource("").toURI().getPath() + entity.getUrl();
            path = path.replace("WEB-INF/classes/", "");
            path = path.replace("file:/", "");
            bufferImg = ImageIO.read(new File(path));
            ImageIO.write(bufferImg, entity.getUrl().substring(entity.getUrl().lastIndexOf(".") + 1, entity.getUrl().length()), byteArrayOut);
            result[0] = byteArrayOut.toByteArray();
            type = entity.getUrl().split("/.")[entity.getUrl().split("/.").length - 1];
        } else {
            result[0] = entity.getData();
            type = PoiPublicUtil.getFileExtendName(entity.getData());
        }
        result[1] = getImageType(type);
        return result;
    }

    /**
     * GET PARAMETER VALUE
     *
     * @param params
     * @param
     * @return
     */
    @SuppressWarnings("rawtypes")
    public static Object getParamsValue(String params, Object object) throws Exception {
        if (params.indexOf(".") != -1) {
            String[] paramsArr = params.split("\\.");
            return getValueDoWhile(object, paramsArr, 0);
        }
        if (object instanceof Map) {
            return ((Map) object).get(params);
        }
        return getMethod(params, object.getClass()).invoke(object, new Object[]{});
    }

    /**
     * Analytical data
     *
     * @return
     */
    public static Object getRealValue(String currentText, Map<String, Object> map) throws Exception {
        String params = "";
        while (currentText.indexOf("{{") != -1) {
            params = currentText.substring(currentText.indexOf("{{") + 2, currentText.indexOf("}}"));
            Object obj = getParamsValue(params.trim(), map);
            // Determine the picture or collection
            if (obj == null) {
                obj = "";
            }
            if (obj instanceof WordImageEntity || obj instanceof List || obj instanceof ExcelListEntity) {
                return obj;
            } else {
                currentText = currentText.replace("{{" + params + "}}", obj.toString());
            }
        }
        return currentText;
    }

    /**
     * 通过遍历过去对象值
     *
     * @param object
     * @param paramsArr
     * @param index
     * @return
     * @throws Exception
     */
    @SuppressWarnings("rawtypes")
    public static Object getValueDoWhile(Object object, String[] paramsArr, int index) throws Exception {
        if (object == null) {
            return "";
        }
        if (object instanceof WordImageEntity) {
            return object;
        }
        if (object instanceof Map) {
            object = ((Map) object).get(paramsArr[index]);
        } else {
            object = getMethod(paramsArr[index], object.getClass()).invoke(object, new Object[]{});
        }
        return (index == paramsArr.length - 1) ? (object == null ? "" : object) : getValueDoWhile(object, paramsArr, ++index);
    }

    /**
     * double to String PREVENT SCIENTIFIC NOTATION
     *
     * @param value
     * @return
     */
    public static String doubleToString(Double value) {
        String temp = value.toString();
        if (temp.contains("E")) {
            BigDecimal bigDecimal = new BigDecimal(temp);
            temp = bigDecimal.toPlainString();
        }
        return ExcelUtil.remove0Suffix(temp);
    }

    /**
     * Determine whether it is a numeric type
     *
     * @param xclass
     * @return
     */
    public static boolean isNumber(String xclass) {
        if (xclass == null) {
            return false;
        }
        String temp = xclass.toLowerCase();
        if (temp.indexOf("int") >= 0 || temp.indexOf("double") >= 0 || temp.indexOf("decimal") >= 0) {
            return true;
        }
        return false;
    }

    /**
     * 统一 KEY ACQUISITION RULES
     *
     * @param key
     * @param targetId
     * @return
     */
    public static String getValueByTargetId(String key, String targetId, String defalut) {
        if (StringUtils.isEmpty(targetId) || key.indexOf("_") < 0) {
            return key;
        }
        String[] arr = key.split(",");
        String[] tempArr;
        for (String str : arr) {
            tempArr = str.split("_");
            if (tempArr == null || tempArr.length < 2) {
                return defalut;
            }
            if (targetId.equals(tempArr[1])) {
                return tempArr[0];
            }
        }
        return defalut;
    }

}
