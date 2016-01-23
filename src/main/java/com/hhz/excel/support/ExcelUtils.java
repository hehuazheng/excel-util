package com.hhz.excel.support;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtils {
	public static Workbook getWorkbook(String fileName) {
		try {
			return WorkbookFactory.create(new File(fileName));
		} catch (InvalidFormatException e) {
			throw new IllegalArgumentException("excel格式不合法", e);
		} catch (IOException e) {
			throw new IllegalArgumentException("生成workbook失败", e);
		}
	}

	public static Workbook getWorkbook(InputStream is) {
		try {
			return WorkbookFactory.create(is);
		} catch (InvalidFormatException e) {
			throw new IllegalArgumentException("excel格式不合法", e);
		} catch (IOException e) {
			throw new IllegalArgumentException("生成workbook失败", e);
		}
	}
}
