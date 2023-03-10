
package com.atviettelsolutions.easypoi.poi.word.parse.excel;

import com.atviettelsolutions.easypoi.poi.common.PoiPublicUtil;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;

/**
 * Process and generate data of type Map into a table
 * 
 * @author caprocute
 * @date 2014年8月9日 下午10:28:46
 */
public final class ExcelMapParse {

	/**
	 * Parse the parameter row to get the parameter list
	 * 
	 * @author caprocute
	 * @param currentRow
	 * @return
	 */
	private static String[] parseCurrentRowGetParams(XWPFTableRow currentRow) {
		List<XWPFTableCell> cells = currentRow.getTableCells();
		String[] params = new String[cells.size()];
		String text;
		for (int i = 0; i < cells.size(); i++) {
			text = cells.get(i).getText();
			params[i] = text == null ? "" : text.trim().replace("{{", "").replace("}}", "");
		}
		return params;
	}

	/**
	 * The next row is parsed and more rows are generated
	 * 
	 * @author caprocute
	 * @param table
	 */
	public static void parseNextRowAndAddRow(XWPFTable table, int index, List<Object> list) throws Exception {
		XWPFTableRow currentRow = table.getRow(index);
		String[] params = parseCurrentRowGetParams(currentRow);
		table.removeRow(index);// Remove this line
		int cellIndex = 0;// The row of the finished object created seems to have an extra cell
		for (Object obj : list) {
			currentRow = table.insertNewTableRow(index++);
			for (cellIndex = 0; cellIndex < currentRow.getTableCells().size(); cellIndex++) {
				String text = PoiPublicUtil.getValueDoWhile(obj, params[cellIndex].split("\\."), 0).toString();
				currentRow.getTableCells().get(cellIndex).setText(text);
			}
			for (; cellIndex < params.length; cellIndex++) {
				currentRow.createCell().setText(PoiPublicUtil.getValueDoWhile(obj, params[cellIndex].split("\\."), 0).toString());
			}
		}

	}

}
