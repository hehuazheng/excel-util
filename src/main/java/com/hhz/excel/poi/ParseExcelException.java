package com.hhz.excel.poi;

public class ParseExcelException extends Exception {
	private static final long serialVersionUID = 8765691838152931080L;

	public ParseExcelException() {
		super();
	}

	public ParseExcelException(String message) {
		super(message);
	}

	public ParseExcelException(String message, Throwable cause) {
		super(message, cause);
	}
}
