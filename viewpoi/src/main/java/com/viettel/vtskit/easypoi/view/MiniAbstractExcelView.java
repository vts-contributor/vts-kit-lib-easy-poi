package com.viettel.vtskit.easypoi.view;

import org.springframework.web.servlet.view.AbstractView;

import javax.servlet.http.HttpServletRequest;

/**
 * Basic abstract Excel View
 * 
 * @author JEECG
 * @date 2015年2月28日 下午1:41:05
 */
public abstract class MiniAbstractExcelView extends AbstractView {

	private static final String CONTENT_TYPE = "application/vnd.ms-excel";

	protected static final String HSSF = ".xls";
	protected static final String XSSF = ".xlsx";

	public MiniAbstractExcelView() {
		setContentType(CONTENT_TYPE);
	}

	protected boolean isIE(HttpServletRequest request) {
		return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0 || request.getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0) ? true : false;
	}

}
