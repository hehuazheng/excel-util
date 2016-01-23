package com.hhz.excel.support;

import java.lang.reflect.Field;
import java.util.List;

import com.google.common.collect.Lists;
import com.hhz.excel.annotation.SheetColumnAttribute;
import com.hhz.excel.poi.FieldWrapper;

public class FieldUtils {
	public static List<FieldWrapper> getFieldWrapperList(Class<?> clazz) {
		List<FieldWrapper> list = Lists.newArrayList();
		Field[] fields = clazz.getDeclaredFields();
		for (int i = 1; i < fields.length; i++) {
			Field field = fields[i - 1];
			SheetColumnAttribute sheetColumn = field
					.getAnnotation(SheetColumnAttribute.class);
			String title = "";
			boolean required = false;
			if (sheetColumn != null) {
				title = sheetColumn.title();
				required = sheetColumn.required();
			}
			list.add(new FieldWrapper(field, title, i, required));
		}
		return list;
	}
}
