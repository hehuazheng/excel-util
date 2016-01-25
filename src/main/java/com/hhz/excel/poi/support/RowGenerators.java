package com.hhz.excel.poi.support;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.hhz.excel.poi.ExcelException;
import com.hhz.excel.poi.FieldWrapper;
import com.hhz.excel.support.FieldUtils;

public class RowGenerators {
	private static RowGenerator DEFAULT_ROW_GENERATOR;

	public static RowGenerator getDefaultGenerator() {
		if (DEFAULT_ROW_GENERATOR == null) {
			synchronized (RowGenerators.class) {
				if (DEFAULT_ROW_GENERATOR == null) {
					DEFAULT_ROW_GENERATOR = new RowGenerator() {
						@Override
						public Row generate(Sheet sheet, Object data)
								throws ExcelException {
							int currentRowNum = sheet.getPhysicalNumberOfRows();
							Row r = sheet.createRow(currentRowNum);
							List<FieldWrapper> list = FieldUtils
									.getFieldWrapperList(data.getClass());
							int colIndex = 1;
							for (FieldWrapper fw : list) {
								Cell cell = r.createCell(colIndex++);
								try {
									Object val = fw.getField().get(data);
									if (val instanceof Number) {
										cell.setCellValue(((Number) val)
												.doubleValue());
									} else if (val instanceof Date) {
										cell.setCellValue((Date) val);
									} else if (val instanceof Boolean) {
										cell.setCellValue((Boolean)val);
									} else if (val instanceof String) {
										cell.setCellValue((String) val);
									} else {
										throw new ExcelException("不支持的类型");
									}
								} catch (ExcelException e) {
									throw e;
								} catch (Exception e) {
									throw new ExcelException("生成行失败", e);
								}
							}
							return r;
						}
					};
				}
			}
		}
		return DEFAULT_ROW_GENERATOR;
	}
}
