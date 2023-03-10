package com.atviettelsolutions.easypoi.poi.excel.html.helper;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * Merge cell help classes
 *
 */
public class MergedRegionHelper {

	private Map<String, Integer[]> mergedCache = new HashMap<String, Integer[]>();

	private Set<String> notNeedCread = new HashSet<String>();

	public MergedRegionHelper(Sheet sheet) {
		getAllMergedRegion(sheet);
	}

	private void getAllMergedRegion(Sheet sheet) {
		int nums = sheet.getNumMergedRegions();
		for (int i = 0; i < nums; i++) {
			handerMergedString(sheet.getMergedRegion(i).formatAsString());
		}
	}

	/**
	 * Based on the merge output, the merge cell matter is handled
	 * 
	 * @param formatAsString
	 */
	private void handerMergedString(String formatAsString) {
		String[] strArr = formatAsString.split(":");
		if (strArr.length == 2) {
			int startCol = strArr[0].charAt(0) - 65;
			if (strArr[0].charAt(1) >= 65) {
				startCol = (startCol + 1) * 26 + (strArr[0].charAt(1) - 65);
			}
			int startRol = Integer.valueOf(strArr[0].substring(strArr[0].charAt(1) >= 65 ? 2 : 1));
			int endCol = strArr[1].charAt(0) - 65;
			if (strArr[1].charAt(1) >= 65) {
				endCol = (endCol + 1) * 26 + (strArr[1].charAt(1) - 65);
			}
			int endRol = Integer.valueOf(strArr[1].substring(strArr[1].charAt(1) >= 65 ? 2 : 1));
			mergedCache.put(startRol + "_" + startCol, new Integer[] { endRol - startRol + 1, endCol - startCol + 1 });
			for (int i = startRol; i <= endRol; i++) {
				for (int j = startCol; j <= endCol; j++) {
					notNeedCread.add(i + "_" + j);
				}
			}
			notNeedCread.remove(startRol + "_" + startCol);
		}

	}

	/**
	 * Is it necessary to create this TD
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	public boolean isNeedCreate(int row, int col) {
		return !notNeedCread.contains(row + "_" + col);
	}

	/**
	 * Is it a merge zone
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	public boolean isMergedRegion(int row, int col) {
		return mergedCache.containsKey(row + "_" + col);
	}

	/**
	 * Gets the merge region
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	public Integer[] getRowAndColSpan(int row, int col) {
		return mergedCache.get(row + "_" + col);
	}

}
