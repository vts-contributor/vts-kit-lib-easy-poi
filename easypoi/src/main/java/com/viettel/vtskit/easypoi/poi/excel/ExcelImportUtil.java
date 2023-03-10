
package com.viettel.vtskit.easypoi.poi.excel;

import com.viettel.vtskit.easypoi.poi.excel.entity.ImportParams;
import com.viettel.vtskit.easypoi.poi.excel.entity.result.ExcelImportResult;
import com.viettel.vtskit.easypoi.poi.excel.imports.ExcelImportServer;
import com.viettel.vtskit.easypoi.poi.excel.imports.sax.SaxReadExcel;
import com.viettel.vtskit.easypoi.poi.excel.imports.sax.parse.ISaxRowRead;
import com.viettel.vtskit.easypoi.poi.handler.inter.IExcelReadRowHanlder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Excel import tool
 * 
 * @author caprocute
 * @version 1.0
 */
@SuppressWarnings({ "unchecked" })
public final class ExcelImportUtil {

	private ExcelImportUtil() {
	}

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelImportUtil.class);

	/**
	 * Excel imports a data source local file without returning validation results Import Segment types Integer, Long, Double, Date, String, Boolean
	 * 
	 * @param file
	 * @param pojoClass
	 * @param params
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> importExcel(File file, Class<?> pojoClass, ImportParams params) {
		FileInputStream in = null;
		List<T> result = null;
		try {
			in = new FileInputStream(file);
			result = new ExcelImportServer().importExcelByIs(in, pojoClass, params).getList();
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				in.close();
			} catch (IOException e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		return result;
	}

	/**
	 * Excel imports the data source IO stream, does not return validation results Import field types Integer, Long, Double, Date, String, Boolean
	 * 
	 * @param file
	 * @param pojoClass
	 * @param params
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> importExcel(InputStream inputstream, Class<?> pojoClass, ImportParams params) throws Exception {
		return new ExcelImportServer().importExcelByIs(inputstream, pojoClass, params).getList();
	}

	/**
	 * Excel imports the data source IO stream, returns the validation results Field types Integer, Long, Double, Date, String, Boolean
	 * 
	 * @param file
	 * @param pojoClass
	 * @param params
	 * @return
	 * @throws Exception
	 */
	public static <T> ExcelImportResult<T> importExcelVerify(InputStream inputstream, Class<?> pojoClass, ImportParams params) throws Exception {
		return new ExcelImportServer().importExcelByIs(inputstream, pojoClass, params);
	}

	/**
	 * Excel imports the data source local file and returns the validation results Field types Integer, Long, Double, Date, String, Boolean
	 * 
	 * @param file
	 * @param pojoClass
	 * @param params
	 * @return
	 * @throws Exception
	 */
	public static <T> ExcelImportResult<T> importExcelVerify(File file, Class<?> pojoClass, ImportParams params) {
		FileInputStream in = null;
		try {
			in = new FileInputStream(file);
			return new ExcelImportServer().importExcelByIs(in, pojoClass, params);
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				in.close();
			} catch (IOException e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		return null;
	}

	/**
	 * Excel uses the SAX parsing method, which is suitable for big data import,
	 * does not support image import data source IO stream, and does not return verification results of the Import field type
	 * Integer,Long,Double,Date,String,Boolean
	 * 
	 * @param inputstream
	 * @param pojoClass
	 * @param params
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> importExcelBySax(InputStream inputstream, Class<?> pojoClass, ImportParams params) {
		return new SaxReadExcel().readExcel(inputstream, pojoClass, params, null, null);
	}

	/**
	 * Excel is suitable for big data import through the SAX parsing method,
	 * does not support image import data source local files, and does not return verification results Import field type
	 * Integer,Long,Double,Date,String,Boolean
	 * 
	 * @param file
	 * @param rowRead
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("rawtypes")
	public static void importExcelBySax(InputStream inputstream, Class<?> pojoClass, ImportParams params, IExcelReadRowHanlder hanlder) {
		new SaxReadExcel().readExcel(inputstream, pojoClass, params, null, hanlder);
	}

	/**
	 * Excel uses the SAX parsing method, which is suitable for big data import,
	 * does not support image import data source IO stream, and does not return verification results of the Import field type
	 * Integer,Long,Double,Date,String,Boolean
	 * 
	 * @param file
	 * @param rowRead
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> importExcelBySax(InputStream inputstream, ISaxRowRead rowRead) {
		return new SaxReadExcel().readExcel(inputstream, null, null, rowRead, null);
	}

}
