package com.hhz.excel.poi;

public class ExcelException extends Exception {
	private static final long serialVersionUID = -9091518508113057861L;

	public ExcelException() {
		super();
	}

	public ExcelException(String message) {
		super(message);
	}

	public ExcelException(String message, Throwable cause) {
		super(message, cause);
	}
}
