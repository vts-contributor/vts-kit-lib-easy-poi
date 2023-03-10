package com.atviettelsolutions.easypoi.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.*;

/**
 * Export service with styles
 * 
 */
public class ExcelExportStylerColorImpl extends AbstractExcelExportStyler implements IExcelExportStyler {

	public ExcelExportStylerColorImpl(Workbook workbook) {
		super.createStyles(workbook);
	}

	@Override
	public CellStyle getHeaderStyle(short headerColor) {
		CellStyle titleStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 24);
		titleStyle.setFont(font);
		titleStyle.setFillForegroundColor(headerColor);
		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return titleStyle;
	}

	@Override
	public CellStyle stringNoneStyle(Workbook workbook, boolean isWarp) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderLeft(BorderStyle.THIN); //
		style.setBorderRight(BorderStyle.THIN); //
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setDataFormat(STRING_FORMAT);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

	@Override
	public CellStyle getTitleStyle(short color) {
		CellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFillForegroundColor(color); //
		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //
		titleStyle.setWrapText(true);
		return titleStyle;
	}

	@Override
	public CellStyle stringSeptailStyle(Workbook workbook, boolean isWarp) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderLeft(BorderStyle.THIN); //
		style.setBorderRight(BorderStyle.THIN); //
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setFillForegroundColor((short) 41); //
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND); //
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setDataFormat(STRING_FORMAT);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

}
