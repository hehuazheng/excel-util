package com.hhz.excel.poi;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.hhz.excel.poi.support.RowGenerator;
import com.hhz.excel.support.FieldUtils;

public class NewExcelGenerator<T> extends ExcelGenerator<T> {
	public NewExcelGenerator(ExcelGeneratorFactory.Builder<T> builder) {
		super(builder.getWorkbook(), builder.getTargetClass());
	}

	@Override
	protected Sheet getProcessSheet() {
		if (workbook.getNumberOfSheets() == 0) {
			workbook.createSheet();
		}
		return workbook.getSheetAt(0);
	}

	@Override
	RowGenerator<T> getRowGenerator() {
		return getDefaultGenerator();
	}

	private RowGenerator<T> rowGenerator;

	public RowGenerator<T> getDefaultGenerator() {
		if (rowGenerator == null) {
			synchronized (NewExcelGenerator.class) {
				if (rowGenerator == null) {
					rowGenerator = new RowGenerator<T>() {
						@Override
						public Row generate(Sheet sheet, T data)
								throws ExcelException {
							int currentRowNum = sheet.getPhysicalNumberOfRows();
							Row r = sheet.createRow(currentRowNum);
							List<FieldWrapper> list = FieldUtils
									.getFieldWrapperList(data.getClass());
							int colIndex = 0;
							for (FieldWrapper fw : list) {
								Cell cell = r.createCell(colIndex++);
								try {
									Object val = fw.getField().get(data);
									if (val == null) {
										continue;
									} else if (val instanceof Number) {
										cell.setCellValue(((Number) val)
												.doubleValue());
									} else if (val instanceof Date) {
										cell.setCellValue((Date) val);
									} else if (val instanceof Boolean) {
										cell.setCellValue((Boolean) val);
									} else if (val instanceof String) {
										cell.setCellValue((String) val);
									} else {
										throw new ExcelException("不支持的类型"
												+ val.toString());
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
		return rowGenerator;
	}
}
