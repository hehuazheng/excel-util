package com.hhz.excel.support;

import java.lang.reflect.Field;
import java.util.Map;

import com.google.common.base.Preconditions;
import com.google.common.collect.Maps;
import com.hhz.excel.annotation.SheetColumnAttribute;
import com.hhz.excel.annotation.SheetAttribute;
import com.hhz.excel.poi.FieldWrapper;

public class AnnotationSheetDefinition extends AbstractSheetDefinition {
	public static int UNPOSSIBLE_INDEX = -1;
	private Map<String, FieldWrapper> titleNameFieldMap = null;

	public AnnotationSheetDefinition(Class<?> clazz) {
		Preconditions.checkArgument(
				clazz.isAnnotationPresent(SheetAttribute.class), clazz
						+ "上未加ExcelModel注解");
		SheetAttribute model = clazz.getAnnotation(SheetAttribute.class);
		super.setTitleRowIndex(model.titleRowIndex());
		titleNameFieldMap = Maps.newHashMap();
		for (Field field : clazz.getDeclaredFields()) {
			SheetColumnAttribute sheetColumn = field
					.getAnnotation(SheetColumnAttribute.class);
			if (sheetColumn != null) {
				titleNameFieldMap.put(sheetColumn.title(), new FieldWrapper(
						field, UNPOSSIBLE_INDEX, sheetColumn.required()));
			}
		}
	}

	public FieldWrapper getFieldByTitleName(String columnName) {
		return titleNameFieldMap.get(columnName.trim());
	}
}
