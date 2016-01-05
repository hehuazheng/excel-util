package com.hhz.excel;

import java.lang.reflect.Field;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.google.common.base.Preconditions;
import com.google.common.collect.Maps;
import com.hhz.excel.annotation.ExcelColumn;
import com.hhz.excel.annotation.ExcelModel;

public class AnnotationExcelDescriptor extends AbstractExcelDescripor {
	private Map<String, Field> titleNameFieldMap = null;

	public AnnotationExcelDescriptor(Class<?> clazz) {
		super();
		Preconditions.checkArgument(
				clazz.isAnnotationPresent(ExcelModel.class), clazz
						+ "上未加ExcelModel注解");
		ExcelModel model = clazz.getAnnotation(ExcelModel.class);
		super.setTitleRowIndex(model.titleRowIndex());
		super.setExtractType(ExtractType.byName);
		titleNameFieldMap = Maps.newHashMap();
		for (Field field : clazz.getDeclaredFields()) {
			ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
			if (excelColumn != null) {
				titleNameFieldMap.put(excelColumn.value(), field);
			}
		}
	}

	public void initFieldMap(Row row) {
		Preconditions.checkNotNull(row, "标题列不能为空");
		int cellCount = row.getPhysicalNumberOfCells();
		for (int i = 1; i <= cellCount; i++) {
			Cell cell = row.getCell(i);
			if (cell != null) {
				String titleName = cell.getStringCellValue().trim();
				Field f = titleNameFieldMap.get(titleName);
				if (f != null) {
					fieldMap.put(i, f);
				}
			}
		}
	}
}
