package com.hhz.excel.support;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.hhz.excel.annotation.SheetAttribute;

public class ExcelUtils {
	public static Workbook getXSSFWorkbook(String fileName) {
		try {
			return WorkbookFactory.create(new File(fileName));
		} catch (InvalidFormatException e) {
			throw new IllegalArgumentException("excel格式不合法", e);
		} catch (IOException e) {
			throw new IllegalArgumentException("生成workbook失败", e);
		}
	}

	public static Workbook getXSSFWorkbook(InputStream is) {
		try {
			return WorkbookFactory.create(is);
		} catch (InvalidFormatException e) {
			throw new IllegalArgumentException("excel格式不合法", e);
		} catch (IOException e) {
			throw new IllegalArgumentException("生成workbook失败", e);
		}
	}

	/**
	 * 获取标题行
	 */
	public static Row getTitleRow(Sheet sheet, Class<?> targetClass) {
		SheetAttribute sheetAttr = targetClass
				.getAnnotation(SheetAttribute.class);
		int titleRowIndex = 0;
		if (sheetAttr != null) {
			titleRowIndex = sheetAttr.titleRowIndex() - 1;
		}
		return sheet.getRow(titleRowIndex);
	}
}
