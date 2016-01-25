package com.hhz.excel.poi;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.google.common.collect.Lists;
import com.hhz.excel.poi.support.FieldNameIndexMapper;
import com.hhz.excel.support.AnnotationSheetDefinition;

public class AnnotationExcelParser<T> extends ExcelParser<T> {
	public AnnotationExcelParser(ExcelParserFactory.Builder<T> builder) {
		super(builder.getWorkbook(), builder.getTargetClass(), builder
				.isMultipleSheetEnabled());
		this.workbook = builder.getWorkbook();
		descriptor = new AnnotationSheetDefinition(targetClass);
	}

	private final AnnotationSheetDefinition descriptor;
	private Map<Integer, FieldWrapper> fieldIndexMap;

	@Override
	protected T convert(Row source) throws ExcelException {
		if (source != null) {
			T obj = newRowModel();
			for (Map.Entry<Integer, FieldWrapper> entry : fieldIndexMap
					.entrySet()) {
				int index = entry.getKey();
				FieldWrapper fw = entry.getValue();
				Field f = fw.getField();
				Cell cell = source.getCell(index);
				try {
					setFieldValue(fw, obj, cell);
				} catch (Exception e) {
					throw new ParseExcelException(f.getName() + "设置值时出错", e);
				}
			}
			return obj;
		}
		return null;
	}

	@Override
	protected List<T> processOneSheet(Sheet sheet) throws ExcelException {
		List<T> list = Lists.newArrayList();
		Row titleRow = sheet.getRow(descriptor.getTitleRowIndex() - 1);
		if (fieldIndexMap == null) {
			fieldIndexMap = FieldNameIndexMapper.toIndexedMap(titleRow,
					descriptor.getFieldWrapperList());
		}
		int rowCount = sheet.getPhysicalNumberOfRows();
		for (int j = descriptor.getTitleRowIndex(); j <= rowCount; j++) {
			Row row = sheet.getRow(j);
			try {
				T t = convert(row);
				if (t != null) {
					list.add(t);
				}
			} catch (Exception e) {
				throw new ExcelException("解析excel异常", e);
			}
		}
		return list;
	}
}
