package com.hhz.excel.support;

import java.lang.reflect.Field;
import java.util.Map;

import com.google.common.base.Preconditions;
import com.google.common.collect.Maps;
import com.hhz.excel.annotation.SheetColumn;
import com.hhz.excel.annotation.SheetDescription;

public class AnnotationSheetDefinition extends AbstractSheetDefinition {
	private Map<String, Field> titleNameFieldMap = null;

	public AnnotationSheetDefinition(Class<?> clazz) {
		super();
		Preconditions.checkArgument(
				clazz.isAnnotationPresent(SheetDescription.class), clazz
						+ "上未加ExcelModel注解");
		SheetDescription model = clazz.getAnnotation(SheetDescription.class);
		super.setTitleRowIndex(model.titleRowIndex());
		titleNameFieldMap = Maps.newHashMap();
		for (Field field : clazz.getDeclaredFields()) {
			SheetColumn sheetColumn = field.getAnnotation(SheetColumn.class);
			if (sheetColumn != null) {
				titleNameFieldMap.put(sheetColumn.value(), field);
			}
		}
	}

	public Field getFieldByTitleName(String columnName) {
		return titleNameFieldMap.get(columnName.trim());
	}
}
