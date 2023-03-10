package com.atviettelsolutions.easypoi.poi.common;


import com.atviettelsolutions.easypoi.poi.exeption.ExcelExportException;

import java.util.Collection;
import java.util.Map;

/**
 * Auto Poi'el Expression Support Utilities
 * 

 */
public final class PoiElUtil {

	public static final String LENGTH = "le:";
	public static final String FOREACH = "fe:";
	public static final String FOREACH_NOT_CREATE = "!fe:";
	public static final String FOREACH_AND_SHIFT = "$fe:";
	public static final String FOREACH_COL        = "#fe:";
	public static final String FOREACH_COL_VALUE  = "v_fe:";
	public static final String START_STR = "{{";
	public static final String END_STR = "}}";
	public static final String WRAP               = "]]";
	public static final String NUMBER_SYMBOL = "n:";
	public static final String FORMAT_DATE = "fd:";
	public static final String FORMAT_NUMBER = "fn:";
	public static final String IF_DELETE = "!if:";
	public static final String EMPTY = "";
	public static final String CONST              = "'";
	public static final String NULL               = "&NULL&";
	public static final String LEFT_BRACKET = "(";
	public static final String RIGHT_BRACKET = ")";
	public static final String DICT_HANDLER       = "dict:";

	private PoiElUtil() {
	}

	/**
	 * Parse strings, support le, fd, fn, !if, trinocular	 *
	 * @param map
	 * @return
	 * @throws Exception
	 */
	public static Object eval(String text, Map<String, Object> map) throws Exception {
		String tempText = new String(text);
		Object obj = innerEval(text, map);
		// If it is not processed and the value exists in the map, process the value
		if (tempText.equals(obj.toString()) && map.containsKey(tempText.split("\\.")[0])) {
			return PoiPublicUtil.getParamsValue(tempText, map);
		}
		return obj;
	}

	/**
	 * Parse strings, support le, fd, fn, !if, trinocular
	 * 
	 * @param map
	 * @return
	 * @throws Exception
	 */
	public static Object innerEval(String text, Map<String, Object> map) throws Exception {
		if (text.indexOf("?") != -1 && text.indexOf(":") != -1) {
			return trinocular(text, map);
		}
		if (text.indexOf(LENGTH) != -1) {
			return length(text, map);
		}
		if (text.indexOf(FORMAT_DATE) != -1) {
			return formatDate(text, map);
		}
		if (text.indexOf(FORMAT_NUMBER) != -1) {
			return formatNumber(text, map);
		}
		if (text.indexOf(IF_DELETE) != -1) {
			return ifDelete(text, map);
		}
		if (text.startsWith("'")) {
			return text.replace("'", "");
		}
		return text;
	}

	/**
	 * Is it to delete the column
	 * 
	 * @param text
	 * @param map
	 * @return
	 * @throws Exception
	 */
	private static Object ifDelete(String text, Map<String, Object> map) throws Exception {
		// Convert multiple spaces into one space
		text = text.replaceAll("\\s{1,}", " ").trim();
		String[] keys = getKey(IF_DELETE, text).split(" ");
		text = text.replace(IF_DELETE, EMPTY);
		return isTrue(keys, map);
	}

	/**
	 * is it true
	 * This method is used in two places
	 * 1. To judge the ternary expression, the expression needs to set spaces {{field == 1? field1:field2 }} or {{field?field1:field2 }}
	 * 2.Take a non-expression (if it is judged to be true, the entire column of the current excel will be killed) {{!if:(field == 1)}} or {{!if:(field)}}
	 *
	 * If the field itself can judge true or false, it will follow the logical processing of len==1
	 * If the field field needs to be combined with other fixed values to judge true or false, then remember to put a space in the expression and then it will split the space to follow the logical processing of len==3
	 * @param keys
	 * @param map
	 * @return
	 * @throws Exception
	 */
	private static Boolean isTrue(String[] keys, Map<String, Object> map) throws Exception {
		if (keys.length == 1) {
			String constant = null;
			if ((constant = isConstant(keys[0])) != null) {
				return Boolean.valueOf(constant);
			}
			return Boolean.valueOf(PoiPublicUtil.getParamsValue(keys[0], map).toString());
		}
		if (keys.length == 3) {
			if(keys[0]==null || keys[2]==null){
				return false;
			}
			Object first = String.valueOf(eval(keys[0], map));
			Object second = String.valueOf(eval(keys[2], map));
			return PoiFunctionUtil.isTrue(first, keys[1], second);
		}
		throw new ExcelExportException("The judgment parameter is wrong");
	}

	/**
	 * Whether the judgment is constant
	 * 
	 * @param param
	 * @return
	 */
	private static String isConstant(String param) {
		if (param.indexOf("'") != -1) {
			return param.replace("'", "");
		}
		return null;
	}

	/**
	 * format number
	 * 
	 * @param text
	 * @param map
	 * @return
	 * @throws Exception
	 */
	private static Object formatNumber(String text, Map<String, Object> map) throws Exception {
		String[] key = getKey(FORMAT_NUMBER, text).split(";");
		text = text.replace(FORMAT_NUMBER, EMPTY);
		return innerEval(replacinnerEvalue(text, PoiFunctionUtil.formatNumber(PoiPublicUtil.getParamsValue(key[0], map), key[1])), map);
	}

	/**
	 * format time
	 * 
	 * @param text
	 * @param map
	 * @return
	 * @throws Exception
	 */
	private static Object formatDate(String text, Map<String, Object> map) throws Exception {
		String[] key = getKey(FORMAT_DATE, text).split(";");
		text = text.replace(FORMAT_DATE, EMPTY);
		return innerEval(replacinnerEvalue(text, PoiFunctionUtil.formatDate(PoiPublicUtil.getParamsValue(key[0], map), key[1])), map);
	}

	/**
	 * Calculate the length of this
	 * 
	 * @param text
	 * @param map
	 * @throws Exception
	 */
	private static Object length(String text, Map<String, Object> map) throws Exception {
		String key = getKey(LENGTH, text);
		text = text.replace(LENGTH, EMPTY);
		Object val = PoiPublicUtil.getParamsValue(key, map);
		return innerEval(replacinnerEvalue(text, PoiFunctionUtil.length(val)), map);
	}

	private static String replacinnerEvalue(String text, Object val) {
		StringBuilder sb = new StringBuilder();
		sb.append(text.substring(0, text.indexOf(LEFT_BRACKET)));
		sb.append(" ");
		sb.append(val);
		sb.append(" ");
		sb.append(text.substring(text.indexOf(RIGHT_BRACKET) + 1, text.length()));
		return sb.toString().trim();
	}

	private static String getKey(String prefix, String text) {
		int leftBracket = 1, rigthBracket = 0, position = 0;
		int index = text.indexOf(prefix) + prefix.length();
		while (text.charAt(index) == " ".charAt(0)) {
			text = text.substring(0, index) + text.substring(index + 1, text.length());
		}
		for (int i = text.indexOf(prefix + LEFT_BRACKET) + prefix.length() + 1; i < text.length(); i++) {
			if (text.charAt(i) == LEFT_BRACKET.charAt(0)) {
				leftBracket++;
			}
			if (text.charAt(i) == RIGHT_BRACKET.charAt(0)) {
				rigthBracket++;
			}
			if (leftBracket == rigthBracket) {
				position = i;
				break;
			}
		}
		return text.substring(text.indexOf(prefix + LEFT_BRACKET) + 1 + prefix.length(), position).trim();
	}

	public static void main(String[] args) {
		System.out.println(getKey(IF_DELETE, "test " + IF_DELETE + " (tom cat)"));
	}

	/**
	 * Trinocular operation
	 * 
	 * @return
	 * @throws Exception
	 */
	private static Object trinocular(String text, Map<String, Object> map) throws Exception {
		// Convert multiple spaces into one space
		text = text.replaceAll("\\s{1,}", " ").trim();
		String testText = text.substring(0, text.indexOf("?"));
		text = text.substring(text.indexOf("?") + 1, text.length()).trim();
		text = innerEval(text, map).toString();
		String[] keys = text.split(":");
		Object first = eval(keys[0].trim(), map);
		Object second = eval(keys[1].trim(), map);
		return isTrue(testText.split(" "), map) ? first : second;
	}

	/**
	 * Parse the string, do not support le, fd, fn,!if, trinocular, get the field prefix of the collection
	 *
	 * @param text
	 * @param map
	 * @return
	 * @throws Exception
	 */
	public static String evalFindName(String text, Map<String, Object> map) throws Exception {
		String[]      keys = text.split("\\.");
		StringBuilder sb   = new StringBuilder().append(keys[0]);
		for (int i = 1; i < keys.length; i++) {
			sb.append(".").append(keys[i]);
			if (eval(sb.toString(), map) instanceof Collection) {
				return sb.toString();
			}
		}
		return null;
	}
}
