
package com.viettel.vtskit.easypoi.poi.word.parse.excel;

import com.viettel.vtskit.easypoi.poi.common.PoiPublicUtil;
import com.viettel.vtskit.easypoi.poi.exeption.word.WordExportException;
import com.viettel.vtskit.easypoi.poi.exeption.word.enmus.WordExportEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import com.viettel.vtskit.easypoi.poi.excel.annotation.ExcelTarget;
import com.viettel.vtskit.easypoi.poi.excel.entity.params.ExcelExportEntity;
import com.viettel.vtskit.easypoi.poi.excel.export.base.ExportBase;
import com.viettel.vtskit.easypoi.poi.word.entity.params.ExcelListEntity;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Resolving Entity Class Objects Reuse annotations
 * 
 * @author caprocute
 */
public class ExcelEntityParse extends ExportBase {

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelEntityParse.class);

	private static void checkExcelParams(ExcelListEntity entity) {
		if (entity.getList() == null || entity.getClazz() == null) {
			throw new WordExportException(WordExportEnum.EXCEL_PARAMS_ERROR);
		}

	}

	private int createCells(int index, Object t, List<ExcelExportEntity> excelParams, XWPFTable table, short rowHeight) throws Exception {
		ExcelExportEntity entity;
		XWPFTableRow row = table.createRow();
		row.setHeight(rowHeight);
		int maxHeight = 1, cellNum = 0;
		for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
			entity = excelParams.get(k);
			if (entity.getList() != null) {
				Collection<?> list = (Collection<?>) entity.getMethod().invoke(t, new Object[] {});
				int listC = 0;
				for (Object obj : list) {
					createListCells(index + listC, cellNum, obj, entity.getList(), table);
					listC++;
				}
				cellNum += entity.getList().size();
				if (list != null && list.size() > maxHeight) {
					maxHeight = list.size();
				}
			} else {
				Object value = getCellValue(entity, t);
				if (entity.getType() == 1) {
					setCellValue(row, value, cellNum++);
				}
			}
		}
		// Merge the cells that need to be merged
		cellNum = 0;
		for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
			entity = excelParams.get(k);
			if (entity.getList() != null) {
				cellNum += entity.getList().size();
			} else if (entity.isNeedMerge()) {
				table.setCellMargins(index, index + maxHeight - 1, cellNum, cellNum);
				cellNum++;
			}
		}
		return maxHeight;
	}

	/**
	 * Create the individual cells after the List
	 * 
	 * @param styles
	 */
	public void createListCells(int index, int cellNum, Object obj, List<ExcelExportEntity> excelParams, XWPFTable table) throws Exception {
		ExcelExportEntity entity;
		XWPFTableRow row;
		if (table.getRow(index) == null) {
			row = table.createRow();
			row.setHeight(getRowHeight(excelParams));
		} else {
			row = table.getRow(index);
		}
		for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
			entity = excelParams.get(k);
			Object value = getCellValue(entity, obj);
			if (entity.getType() == 1) {
				setCellValue(row, value, cellNum++);
			}
		}
	}

	/**
	 * Get the header data
	 * 
	 * @param table
	 * @param index
	 * @return
	 */
	private Map<String, Integer> getTitleMap(XWPFTable table, int index, int headRows) {
		if (index < headRows) {
			throw new WordExportException(WordExportEnum.EXCEL_NO_HEAD);
		}
		Map<String, Integer> map = new HashMap<String, Integer>();
		String text;
		for (int j = 0; j < headRows; j++) {
			List<XWPFTableCell> cells = table.getRow(index - j - 1).getTableCells();
			for (int i = 0; i < cells.size(); i++) {
				text = cells.get(i).getText();
				if (StringUtils.isEmpty(text)) {
					throw new WordExportException(WordExportEnum.EXCEL_HEAD_HAVA_NULL);
				}
				map.put(text, i);
			}
		}
		return map;
	}

	/**
	 * Parse the previous line and generate more rows
	 * 
	 * @param table

	 */
	public void parseNextRowAndAddRow(XWPFTable table, int index, ExcelListEntity entity) {
		checkExcelParams(entity);
		// Get the header data
		Map<String, Integer> titlemap = getTitleMap(table, index, entity.getHeadRows());
		try {
			// Get all fields
			Field fileds[] = PoiPublicUtil.getClassFields(entity.getClazz());
			ExcelTarget etarget = entity.getClazz().getAnnotation(ExcelTarget.class);
			String targetId = null;
			if (etarget != null) {
				targetId = etarget.value();
			}
			// Gets the exported data for an entity object
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			getAllExcelField(null, targetId, fileds, excelParams, entity.getClazz(), null);
			// Filter sorting based on table headers
			sortAndFilterExportField(excelParams, titlemap);
			short rowHeight = getRowHeight(excelParams);
			Iterator<?> its = entity.getList().iterator();
			while (its.hasNext()) {
				Object t = its.next();
				index += createCells(index, t, excelParams, table, rowHeight);
			}
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		}
	}

	private void setCellValue(XWPFTableRow row, Object value, int cellNum) {
		if (row.getCell(cellNum++) != null) {
			row.getCell(cellNum - 1).setText(value == null ? "" : value.toString());
		} else {
			row.createCell().setText(value == null ? "" : value.toString());
		}
	}

	/**
	 * Sort and plug the exported sequence
	 * 
	 * @param excelParams
	 * @param titlemap
	 * @return
	 */
	private void sortAndFilterExportField(List<ExcelExportEntity> excelParams, Map<String, Integer> titlemap) {
		for (int i = excelParams.size() - 1; i >= 0; i--) {
			if (excelParams.get(i).getList() != null && excelParams.get(i).getList().size() > 0) {
				sortAndFilterExportField(excelParams.get(i).getList(), titlemap);
				if (excelParams.get(i).getList().size() == 0) {
					excelParams.remove(i);
				} else {
					excelParams.get(i).setOrderNum(i);
				}
			} else {
				if (titlemap.containsKey(excelParams.get(i).getName())) {
					excelParams.get(i).setOrderNum(i);
				} else {
					excelParams.remove(i);
				}
			}
		}
		sortAllParams(excelParams);
	}

}
