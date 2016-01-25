package com.hhz.excel.support;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

import com.google.common.collect.Lists;
import com.hhz.excel.annotation.SheetAttribute;
import com.hhz.excel.annotation.SheetColumnAttribute;
import com.hhz.excel.poi.FieldWrapper;

public class AnnotationSheetDefinition extends AbstractSheetDefinition {
	private Map<String, FieldWrapper> titleNameFieldMap = null;

	public AnnotationSheetDefinition(Class<?> clazz) {
		SheetAttribute attr = clazz.getAnnotation(SheetAttribute.class);
		if (attr != null) {
			super.setTitleRowIndex(attr.titleRowIndex());
		}
		this.fieldWrapperList = getFieldWrapperList(clazz);
	}

	public FieldWrapper getFieldByTitleName(String columnName) {
		return titleNameFieldMap.get(columnName.trim());
	}

	public static List<FieldWrapper> getFieldWrapperList(Class<?> clazz) {
		List<FieldWrapper> list = Lists.newArrayList();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 1; i < fields.length; i++) {
			Field field = fields[i - 1];
			SheetColumnAttribute sheetColumn = field
					.getAnnotation(SheetColumnAttribute.class);
			if (sheetColumn != null) {
				list.add(new FieldWrapper(field, sheetColumn.title(), i,
						sheetColumn.required()));
			}
		}
		return list;
	}

}
