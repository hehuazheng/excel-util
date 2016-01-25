package com.hhz.excel.poi.support;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.hhz.excel.poi.ExcelException;

public interface RowGenerator<T> {
	Row generate(Sheet sheet, T rowData) throws ExcelException;
}
