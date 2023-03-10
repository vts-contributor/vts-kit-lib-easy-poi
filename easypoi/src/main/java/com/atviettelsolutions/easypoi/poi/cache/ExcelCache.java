package com.atviettelsolutions.easypoi.poi.cache;

import com.atviettelsolutions.easypoi.poi.cache.manager.POICacheManager;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
public final class ExcelCache {

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelCache.class);

	public static Workbook getWorkbook(String url, Integer[] sheetNums, boolean needAll) {
		InputStream is = null;
		List<Integer> sheetList = Arrays.asList(sheetNums);
		try {
			is = POICacheManager.getFile(url);
			Workbook wb = WorkbookFactory.create(is);
			// delete other sheets
			if (!needAll) {
				for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
					if (!sheetList.contains(i)) {
						wb.removeSheetAt(i);
					}
				}
			}
			return wb;
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				is.close();
			} catch (Exception e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		return null;
	}
	public static Workbook getWorkbookByTemplate(String url, Integer[] sheetNums, boolean needAll) {
		List<Integer> sheetList = Arrays.asList(sheetNums);
		InputStream fis = null;
		try {
			//ClassPathResource  resource = new ClassPathResource(url);
			fis = new FileInputStream(url);
			LOGGER.info("  >>>  POI 3 UPGRADE TO 4 COMPATIBLE RETROFIT WORK, url="+url);
			//fis = resource.getInputStream();
			Workbook wb = WorkbookFactory.create(fis);
			// delete other sheets
			if (!needAll) {
				for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
					if (!sheetList.contains(i)) {
						wb.removeSheetAt(i);
					}
				}
			}
			return wb;
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				fis.close();
			} catch (Exception e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		return null;
	}
}
