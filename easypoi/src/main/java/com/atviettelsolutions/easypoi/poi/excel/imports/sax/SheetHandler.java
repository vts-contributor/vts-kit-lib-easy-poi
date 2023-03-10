
package com.atviettelsolutions.easypoi.poi.excel.imports.sax;

import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import com.atviettelsolutions.easypoi.poi.excel.entity.enmus.CellValueType;
import com.atviettelsolutions.easypoi.poi.excel.entity.sax.SaxReadCellEntity;
import com.atviettelsolutions.easypoi.poi.excel.imports.sax.parse.ISaxRowRead;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

/**
 * Callback interface
 * 
 * @author caprocute
 */
public class SheetHandler extends DefaultHandler {

	private SharedStringsTable sst;
	private String lastContents;

	private int curRow = 0;
	private int curCol = 0;

	private CellValueType type;

	private ISaxRowRead read;

	// The container in which row records are stored
	private List<SaxReadCellEntity> rowlist = Lists.newArrayList();

	public SheetHandler(SharedStringsTable sst, ISaxRowRead rowRead) {
		this.sst = sst;
		this.read = rowRead;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// empty
		lastContents = "";
		// c => cell
		if ("c".equals(name)) {
			// If the next element is the index of the SST, mark nextIsString as true
			String cellType = attributes.getValue("t");
			if ("s".equals(cellType)) {
				type = CellValueType.String;
				return;
			}
			// Date format
			cellType = attributes.getValue("s");
			if ("1".equals(cellType)) {
				type = CellValueType.Date;
			} else if ("2".equals(cellType)) {
				type = CellValueType.Number;
			}
		} else if ("t".equals(name)) {// When the element is t
			type = CellValueType.TElement;
		}

	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {

		// According to the SST's index value, the string to be stored into the cell's real to store
		// The characters() method may be called multiple times
		if (CellValueType.String.equals(type)) {
			try {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			} catch (Exception e) {

			}
		}
		// The t element also contains strings
		if (CellValueType.TElement.equals(type)) {
			String value = lastContents.trim();
			rowlist.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
			curCol++;
			type = CellValueType.None;
			// v = > the value of the cell, if the cell is a string, the value of the v tag is the index of that string in the SST
			// Add the contents of the cell to the rowlist, removing the white space before and after the string
		} else if ("v".equals(name)) {
			String value = lastContents.trim();
			value = value.equals("") ? " " : value;
			if (CellValueType.Date.equals(type)) {
				Date date = HSSFDateUtil.getJavaDate(Double.valueOf(value));
				rowlist.add(curCol, new SaxReadCellEntity(CellValueType.Date, date));
			} else if (CellValueType.Number.equals(type)) {
				BigDecimal bd = new BigDecimal(value);
				rowlist.add(curCol, new SaxReadCellEntity(CellValueType.Number, bd));
			} else if (CellValueType.String.equals(type)) {
				rowlist.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
			}
			curCol++;
		} else if (name.equals("row")) {//If the tag name is row, this indicates that the end of the line is reached, call the optRows() method
			read.parse(curRow, rowlist);
			rowlist.clear();
			curRow++;
			curCol = 0;
		}

	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		// Get the value of the cell contents
		lastContents += new String(ch, start, length);
	}

}
