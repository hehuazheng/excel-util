package com.hhz.excel.poi;

import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.hhz.excel.poi.support.FieldNameIndexMapper;
import com.hhz.excel.poi.support.RowGenerator;
import com.hhz.excel.support.ExcelUtils;
import com.hhz.excel.support.FieldUtils;

public class TemplateExcelGenerator<T> extends ExcelGenerator<T> {
	private MyRowGenerator<T> myRowGenerator;

	public TemplateExcelGenerator(ExcelGeneratorFactory.Builder<T> builder) {
		super(builder.getWorkbook(), builder.getTargetClass());
		myRowGenerator = new MyRowGenerator<T>(getProcessSheet(), targetClass);
	}

	@Override
	protected Sheet getProcessSheet() {
		return workbook.getSheetAt(0);
	}

	@Override
	RowGenerator getRowGenerator() {
		return myRowGenerator;
	}

	private static class MyRowGenerator<D> implements RowGenerator {
		Map<Integer, FieldWrapper> indexFieldMap;
		private int columnCount = 0;

		public MyRowGenerator(Sheet sheet, Class<D> targetClass) {
			Row titleRow = ExcelUtils.getTitleRow(sheet, targetClass);
			columnCount = titleRow.getPhysicalNumberOfCells();
			List<FieldWrapper> fieldWrapperList = FieldUtils
					.getFieldWrapperList(targetClass);
			indexFieldMap = FieldNameIndexMapper.toIndexedMap(titleRow,
					fieldWrapperList);
		}

		@Override
		public Row generate(Sheet sheet, Object rowData) throws ExcelException {
			int currentRowNum = sheet.getPhysicalNumberOfRows();
			Row r = sheet.createRow(currentRowNum);
			int colIndex = 1;
			for (int i = 0; i < columnCount; i++) {
				Cell cell = r.createCell(colIndex++);
				FieldWrapper fw = indexFieldMap.get(i);
				if (fw != null) {
					try {
						Object val = fw.getField().get(rowData);
						if (val == null) {
							continue;
						} else if (val instanceof Number) {
							cell.setCellValue(((Number) val).doubleValue());
						} else if (val instanceof Date) {
							cell.setCellValue((Date) val);
						} else if (val instanceof Boolean) {
							cell.setCellValue((Boolean) val);
						} else if (val instanceof String) {
							cell.setCellValue((String) val);
						} else {
							throw new ExcelException("不支持的类型" + val.toString());
						}
					} catch (ExcelException e) {
						throw e;
					} catch (Exception e) {
						throw new ExcelException("生成行失败", e);
					}
				}
			}
			return r;
		}
	}
}
