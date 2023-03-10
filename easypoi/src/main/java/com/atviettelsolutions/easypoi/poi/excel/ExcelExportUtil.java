
package com.atviettelsolutions.easypoi.poi.excel;

import com.atviettelsolutions.easypoi.poi.excel.export.ExcelBatchExportServer;
import com.atviettelsolutions.easypoi.poi.excel.export.ExcelExportServer;
import com.atviettelsolutions.easypoi.poi.handler.inter.IExcelExportServer;
import com.atviettelsolutions.easypoi.poi.handler.inter.IWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.atviettelsolutions.easypoi.poi.excel.entity.ExportParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.TemplateExportParams;
import com.atviettelsolutions.easypoi.poi.excel.entity.enmus.ExcelType;
import com.atviettelsolutions.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.atviettelsolutions.easypoi.poi.excel.export.template.ExcelExportOfTemplateUtil;

import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * excel Export the tool class
 * 
 * @author caprocute
 * @version 1.0
 */
public final class ExcelExportUtil {
	//Single sheet maximum
	public static       int    USE_SXSSF_LIMIT = 100000;
	private ExcelExportUtil() {
	}

	/**
	 * Create the corresponding Excel according to the Entity
	 * 
	 * @param entity
	 *
	 * @param pojoClass
	 *
	 * @param dataSet
	 *
	 * @param exportFields
	 *
	 */
	public static Workbook exportExcel(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet,String[] exportFields) {
		Workbook workbook;
		if (ExcelType.HSSF.equals(entity.getType())) {
			workbook = new HSSFWorkbook();
		} else if (dataSet.size() < 1000) {
			workbook = new XSSFWorkbook();
		} else {
			workbook = new SXSSFWorkbook();
		}
		new ExcelExportServer().createSheet(workbook, entity, pojoClass, dataSet,exportFields);
		return workbook;
	}


	/**
	 * Create the corresponding Excel according to the Entity
	 *
	 * @param entity
	 *
	 * @param pojoClass
	 *
	 * @param dataSet
	 *
	 */
	public static Workbook exportExcel(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet) {
		Workbook workbook;
		if (ExcelType.HSSF.equals(entity.getType())) {
			workbook = new HSSFWorkbook();
		} else if (dataSet.size() < 1000) {
			workbook = new XSSFWorkbook();
		} else {
			workbook = new SXSSFWorkbook();
		}
		new ExcelExportServer().createSheet(workbook, entity, pojoClass, dataSet,null);
		return workbook;
	}

	/**
	 * Create an Excel based on the Map
	 * 
	 * @param entity
	 *
	 *
	 * @param dataSet
	 *
	 */
	public static Workbook exportExcel(ExportParams entity, List<ExcelExportEntity> entityList, Collection<? extends Map<?, ?>> dataSet) {
		Workbook workbook;
		if (ExcelType.HSSF.equals(entity.getType())) {
			workbook = new HSSFWorkbook();
		} else if (dataSet.size() < 1000) {
			workbook = new XSSFWorkbook();
		} else {
			workbook = new SXSSFWorkbook();
		}
		new ExcelExportServer().createSheetForMap(workbook, entity, entityList, dataSet);
		return workbook;
	}

	/**
	 * One Excel creates multiple sheets
	 * 
	 * @param list
	 *            Multiple Map key titles correspond to tables, and Title key entities correspond to table corresponding entity key data Collection data
	 * @return
	 */
	public static Workbook exportExcel(List<Map<String, Object>> list, ExcelType type) {
		Workbook workbook;
		if (ExcelType.HSSF.equals(type)) {
			workbook = new HSSFWorkbook();
		} else {
			workbook = new XSSFWorkbook();
		}
		for (Map<String, Object> map : list) {
			ExcelExportServer server = new ExcelExportServer();
			server.createSheet(workbook, (ExportParams) map.get("title"), (Class<?>) map.get("entity"), (Collection<?>) map.get("data"),null);
		}
		return workbook;
	}

	/**
	 * The export file is parsed by the template, this is not recommended, it is recommended to perform all processing through the template
	 * 
	 * @param params
	 * @param pojoClass
	 * @param dataSet
	 * @param map
	 * @return
	 */
	public static Workbook exportExcel(TemplateExportParams params, Class<?> pojoClass, Collection<?> dataSet, Map<String, Object> map) {
		return new ExcelExportOfTemplateUtil().createExcleByTemplate(params, pojoClass, dataSet, map);
	}

	/**
	 * The export file parses only templates and no collections through templates
	 * 
	 * @param params
	 * @param map
	 * @return
	 */
	public static Workbook exportExcel(TemplateExportParams params, Map<String, Object> map) {
		return new ExcelExportOfTemplateUtil().createExcleByTemplate(params, null, null, map);
	}


	/**
	 * Export of large data volumes
	 *
	 * @param entity    Table header property
	 * @param pojoClass Excel object Class
	 * @return ExcelBatchExportServer batch_service
	 */
	public static IWriter<Workbook> exportBigExcel(ExportParams entity, Class<?> pojoClass) {
		ExcelBatchExportServer batchServer = new ExcelBatchExportServer();
		batchServer.init(entity, pojoClass);
		return batchServer;
	}

	/**
	 * Export of large data volumes
	 *
	 * @param entity
	 * @param excelParams
	 * @return ExcelBatchExportServer Batch service
	 */
	public static IWriter<Workbook> exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams) {
		ExcelBatchExportServer batchServer = new ExcelBatchExportServer();
		batchServer.init(entity, excelParams);
		return batchServer;
	}

	/**
	 * Export of large data volumes
	 *
	 * @param entity
	 * @param pojoClass
	 * @param server
	 * @param queryParams
	 * @return Workbook
	 */
	public static Workbook exportBigExcel(ExportParams entity, Class<?> pojoClass,
                                          IExcelExportServer server, Object queryParams) {
		ExcelBatchExportServer batchServer = new ExcelBatchExportServer();
		batchServer.init(entity, pojoClass);
		return batchServer.exportBigExcel(server, queryParams);
	}

	/**
	 * Export of large data volumes
	 * @param entity
	 * @param excelParams
	 * @param server
	 * @param queryParams
	 * @return Workbook
	 */
	public static Workbook exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams,
										  IExcelExportServer server, Object queryParams) {
		ExcelBatchExportServer batchServer = new ExcelBatchExportServer();
		batchServer.init(entity, excelParams);
		return batchServer.exportBigExcel(server, queryParams);
	}
}
