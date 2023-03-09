package com.viettel.vtskit.easypoi.poi.exeption.word;

import com.viettel.vtskit.easypoi.poi.exeption.word.enmus.WordExportEnum;

public class WordExportException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public WordExportException() {
		super();
	}

	public WordExportException(String msg) {
		super(msg);
	}

	public WordExportException(WordExportEnum exception) {
		super(exception.getMsg());
	}

}
