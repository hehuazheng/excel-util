package com.hhz.excel;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.hhz.common.Converter;
import com.hhz.excel.util.CellUtil;

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
			// int sheetCount = workbook.getNumberOfSheets();
			// for (int i = 0; i < sheetCount; i++) {
			// list.addAll(processOneSheet(workbook.getSheetAt(i)));
			list.addAll(processOneSheet(workbook.getSheetAt(0)));
			// }
		}
		return list;
	}

	private List<T> processOneSheet(Sheet sheet) {
		List<T> list = Lists.newArrayList();
		Row titleRow = sheet.getRow(descriptor.getTitleRowIndex());
		this.descriptor.initFieldMap(titleRow);
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

	private T convert(Row source) {
		if (source != null) {
			try {
				T r = targetClass.newInstance();
				for (int i = 1; i <= source.getPhysicalNumberOfCells(); i++) {
					Field f = this.descriptor.getFieldMap().get(i).getField();
					if (f != null) {
						Cell cell = source.getCell(i);
						setField(f, r, cell);
					}
				}
				return r;
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return null;
	}

	private static Converter<String, Integer> INTEGER_CONVERTER = new Converter<String, Integer>() {
		@Override
		public Integer convert(String source) {
			if (source != null && !"".equals(source)) {
				return Integer.parseInt(source);
			}
			return null;
		}
	};
	private static Converter<String, Double> DOUBLE_CONVERTER = new Converter<String, Double>() {
		@Override
		public Double convert(String source) {
			if (source != null && !"".equals(source)) {
				return Double.parseDouble(source);
			}
			return null;
		}
	};
	private static Converter<String, String> STRING_CONVERTER = new Converter<String, String>() {
		@Override
		public String convert(String source) {
			return source;
		}
	};

	private Map<Class<?>, Converter<String, ?>> defaultConverterMap = Maps
			.newHashMap();
	{
		defaultConverterMap.put(int.class, INTEGER_CONVERTER);
		defaultConverterMap.put(Integer.class, INTEGER_CONVERTER);
		defaultConverterMap.put(Double.class, DOUBLE_CONVERTER);
		defaultConverterMap.put(double.class, DOUBLE_CONVERTER);
		defaultConverterMap.put(String.class, STRING_CONVERTER);
	}

	private Map<Class<?>, Converter<String, ?>> getDefaultConverters() {
		return defaultConverterMap;
	}

	private void setField(Field f, Object obj, Cell cell)
			throws ParseExcelException {
		Converter<String, ?> converter = getDefaultConverters()
				.get(f.getType());
		if (converter == null) {
			throw new ParseExcelException("不支持的转换类型");
		}
		String cellStringValue = CellUtil.readCellValueToString(cell);
		try {
			f.set(obj, converter.convert(cellStringValue));
		} catch (Exception e) {
			throw new ParseExcelException(f.getName() + "设置值时出错"
					+ cellStringValue, e);
		}
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
