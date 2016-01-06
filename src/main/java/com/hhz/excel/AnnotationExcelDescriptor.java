package com.hhz.excel;

import java.lang.reflect.Field;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.google.common.base.Preconditions;
import com.google.common.collect.Maps;
import com.hhz.excel.annotation.SheetColumn;
import com.hhz.excel.annotation.SheetModel;

public class AnnotationExcelDescriptor extends AbstractExcelDescripor {
	private Map<String, Field> titleNameFieldMap = null;

	public AnnotationExcelDescriptor(Class<?> clazz) {
		super();
		Preconditions.checkArgument(
				clazz.isAnnotationPresent(SheetModel.class), clazz
						+ "上未加ExcelModel注解");
		SheetModel model = clazz.getAnnotation(SheetModel.class);
		super.setTitleRowIndex(model.titleRowIndex());
		titleNameFieldMap = Maps.newHashMap();
		for (Field field : clazz.getDeclaredFields()) {
			SheetColumn sheetColumn = field.getAnnotation(SheetColumn.class);
			if (sheetColumn != null) {
				titleNameFieldMap.put(sheetColumn.value(), field);
			}
		}
	}

	public void initFieldMap(Row row) {
		if (getFieldWrapperList().isEmpty()) {
			Preconditions.checkNotNull(row, "标题列不能为空");
			int cellCount = row.getPhysicalNumberOfCells();
			for (int i = 1; i <= cellCount; i++) {
				Cell cell = row.getCell(i);
				if (cell != null) {
					String titleName = cell.getStringCellValue().trim();
					Field f = titleNameFieldMap.get(titleName);
					if (f != null) {
						addFieldWrapper(new FieldWrapper(f, i));
					}
				}
			}
		}
	}
}
