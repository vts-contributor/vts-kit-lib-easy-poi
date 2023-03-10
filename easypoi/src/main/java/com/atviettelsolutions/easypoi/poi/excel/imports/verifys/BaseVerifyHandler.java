
package com.atviettelsolutions.easypoi.poi.excel.imports.verifys;


import com.atviettelsolutions.easypoi.poi.excel.entity.result.ExcelVerifyHanlderResult;

import java.util.regex.Pattern;

/**
 * The underlying validation utility class
 * 
 * @author caprocute
 */
public class BaseVerifyHandler {

	private static String NOT_NULL = "Null is not allowed";
	private static String IS_MOBILE = "Not a mobile phone number";
	private static String IS_TEL = "Not a phone number";
	private static String IS_EMAIL = "Not an email address";
	private static String MIN_LENGHT = "Less than the specified length";
	private static String MAX_LENGHT = "Exceeding the prescribed length";

	private static Pattern mobilePattern = Pattern.compile("^[1][3,4,5,8,7][0-9]{9}$");

	private static Pattern telPattern = Pattern.compile("^([0][1-9]{2,3}-)?[0-9]{5,10}$");

	private static Pattern emailPattern = Pattern.compile("^([a-zA-Z0-9]+[_|\\_|\\.]?)*[a-zA-Z0-9]+@([a-zA-Z0-9]+[_|\\_|\\.]?)*[a-zA-Z0-9]+\\.[a-zA-Z]{2,3}$");

	/**
	 * email check
	 * 
	 * @param name
	 * @param val
	 * @return
	 */
	public static ExcelVerifyHanlderResult isEmail(String name, Object val) {
		if (!emailPattern.matcher(String.valueOf(val)).matches()) {
			return new ExcelVerifyHanlderResult(false, name + IS_EMAIL);
		}
		return new ExcelVerifyHanlderResult(true);
	}

	/**
	 *
	 * 
	 * @param name
	 * @param val
	 * @return
	 */
	public static ExcelVerifyHanlderResult isMobile(String name, Object val) {
		if (!mobilePattern.matcher(String.valueOf(val)).matches()) {
			return new ExcelVerifyHanlderResult(false, name + IS_MOBILE);
		}
		return new ExcelVerifyHanlderResult(true);
	}

	/**
	 *
	 * 
	 * @param name
	 * @param val
	 * @return
	 */
	public static ExcelVerifyHanlderResult isTel(String name, Object val) {
		if (!telPattern.matcher(String.valueOf(val)).matches()) {
			return new ExcelVerifyHanlderResult(false, name + IS_TEL);
		}
		return new ExcelVerifyHanlderResult(true);
	}

	/**
	 * Maximum length check
	 * 
	 * @param name
	 * @param val
	 * @return
	 */
	public static ExcelVerifyHanlderResult maxLength(String name, Object val, int maxLength) {
		if (notNull(name, val).isSuccess() && String.valueOf(val).length() > maxLength) {
			return new ExcelVerifyHanlderResult(false, name + MAX_LENGHT);
		}
		return new ExcelVerifyHanlderResult(true);
	}

	/**
	 * ExcelVerifyHanlderResult
	 * 
	 * @param name
	 * @param val
	 * @param minLength
	 * @return
	 */
	public static ExcelVerifyHanlderResult minLength(String name, Object val, int minLength) {
		if (notNull(name, val).isSuccess() && String.valueOf(val).length() < minLength) {
			return new ExcelVerifyHanlderResult(false, name + MIN_LENGHT);
		}
		return new ExcelVerifyHanlderResult(true);
	}

	/**
	 * ExcelVerifyHanlderResult
	 * 
	 * @param name
	 * @param val
	 * @return
	 */
	public static ExcelVerifyHanlderResult notNull(String name, Object val) {
		if (val == null || val.toString().equals("")) {
			return new ExcelVerifyHanlderResult(false, name + NOT_NULL);
		}
		return new ExcelVerifyHanlderResult(true);
	}

	/**
	 * ExcelVerifyHanlderResult
	 * 
	 * @param name
	 * @param val
	 * @param regex
	 * @param regexTip
	 * @return
	 */
	public static ExcelVerifyHanlderResult regex(String name, Object val, String regex, String regexTip) {
		Pattern pattern = Pattern.compile(regex);
		if (!pattern.matcher(String.valueOf(val)).matches()) {
			return new ExcelVerifyHanlderResult(false, name + regexTip);
		}
		return new ExcelVerifyHanlderResult(true);
	}

}
