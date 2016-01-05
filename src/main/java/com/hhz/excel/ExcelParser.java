package com.hhz.excel;

import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;
import com.google.common.collect.Lists;

public class ExcelParser<T> {
	private final AnnotationExcelDescriptor descriptor;
	private final Class<T> targetClass;

	private ExcelParser(Class<T> targetClass) {
		this.targetClass = targetClass;
		this.descriptor = new AnnotationExcelDescriptor(targetClass);
	}

	public List<T> parse(Workbook workbook) {
		List<T> list = Lists.newArrayList();
		if (workbook != null) {
//			int sheetCount = workbook.getNumberOfSheets();
//			for (int i = 0; i < sheetCount; i++) {
//				list.addAll(processOneSheet(workbook.getSheetAt(i)));
				list.addAll(processOneSheet(workbook.getSheetAt(0)));
//			}
		}
		return list;
	}

	private List<T> processOneSheet(Sheet sheet) {
		List<T> list = Lists.newArrayList();
		this.descriptor
				.initFieldMap(sheet.getRow(descriptor.getTitleRowIndex()));
		int rowCount = sheet.getPhysicalNumberOfRows();
		for (int i = descriptor.getTitleRowIndex() + 1; i <= rowCount; i++) {
			Row row = sheet.getRow(i);
			T t = convert(row);
			if (t != null) {
				list.add(t);
			}
		}
		return list;
	}

	public T convert(Row source) {
		if (source != null) {
			try {
				T r = targetClass.newInstance();
				for (int i = 1; i <= source.getPhysicalNumberOfCells(); i++) {
					Field f = this.descriptor.getFieldMap().get(i);
					if (f != null) {
						Cell cell = source.getCell(i);
						f.setAccessible(true);
						try {
							f.set(r, cell.getStringCellValue());
						} catch (IllegalArgumentException
								| IllegalAccessException e) {
							e.printStackTrace();
						}
					}
				}
				return r;
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return null;
	}

	public static class ExcelParserBuilder<T> {
		private Class<T> targetClass;

		private ExcelParserBuilder(Class<T> targetClass) {
			this.targetClass = targetClass;
		}

		public static <T> ExcelParserBuilder<T> create(Class<T> targetClass) {
			return new ExcelParserBuilder<T>(targetClass);
		}

		public ExcelParser<T> build() {
			Preconditions.checkNotNull(targetClass, "targetClass不能为空");
			return new ExcelParser<T>(targetClass);
		}
	}
}
