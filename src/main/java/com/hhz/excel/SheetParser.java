package com.hhz.excel;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Preconditions;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.hhz.common.Converter;
import com.hhz.excel.util.CellUtil;

public class SheetParser<T> {
	private final AnnotationExcelDescriptor descriptor;
	private final Class<T> targetClass;
	private Map<Class<?>, Converter<String, ?>> converterMap;
	private Sheet sheet;

	private SheetParser(Class<T> targetClass,
			Map<Class<?>, Converter<String, ?>> converterMap, Sheet sheet) {
		this.targetClass = targetClass;
		this.descriptor = new AnnotationExcelDescriptor(targetClass);
		this.converterMap = converterMap;
		this.sheet = sheet;
	}

	public List<T> toList() {
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

	private void setField(Field f, Object obj, Cell cell)
			throws ParseExcelException {
		Converter<String, ?> converter = converterMap.get(f.getType());
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

	public static class SheetParserBuilder<T> {
		private static final Logger LOGGER = LoggerFactory
				.getLogger(SheetParserBuilder.class);
		private Class<T> targetClass;
		private Map<Class<?>, Converter<String, ?>> onverterMap;
		private boolean useDefaultConverterMap = true;
		private Workbook workbook;

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

		private static Map<Class<?>, Converter<String, ?>> defaultConverterMap = Maps
				.newHashMap();
		{
			defaultConverterMap.put(int.class, INTEGER_CONVERTER);
			defaultConverterMap.put(Integer.class, INTEGER_CONVERTER);
			defaultConverterMap.put(Double.class, DOUBLE_CONVERTER);
			defaultConverterMap.put(double.class, DOUBLE_CONVERTER);
			defaultConverterMap.put(String.class, STRING_CONVERTER);
		}

		private SheetParserBuilder(Class<T> targetClass) {
			this.targetClass = targetClass;
		}

		public SheetParserBuilder<T> addCustomConverter(Class<?> clazz,
				Converter<String, ?> converter) {
			if (onverterMap == null) {
				onverterMap = Maps.newHashMap();
			}
			return this;
		}

		/**
		 * 不使用默认的转换器
		 */
		public SheetParserBuilder<T> notUseDefaultConverter() {
			this.useDefaultConverterMap = false;
			return this;
		}

		public SheetParserBuilder<T> setWorkbook(Workbook workbook) {
			this.workbook = workbook;
			return this;
		}

		public SheetParserBuilder<T> setFilePath(String filePath) {
			try {
				this.workbook = WorkbookFactory.create(new FileInputStream(
						filePath));
			} catch (Exception e) {
				LOGGER.error("生成workbook异常 " + filePath, e);
			}
			return this;
		}

		public static <T> SheetParserBuilder<T> create(Class<T> targetClass) {
			return new SheetParserBuilder<T>(targetClass);
		}

		public SheetParser<T> build() {
			Preconditions.checkNotNull(targetClass, "targetClass不能为空");
			if (onverterMap == null) {
				onverterMap = Maps.newHashMap();
			}
			if (useDefaultConverterMap) {
				onverterMap.put(int.class, INTEGER_CONVERTER);
				onverterMap.put(Integer.class, INTEGER_CONVERTER);
				onverterMap.put(Double.class, DOUBLE_CONVERTER);
				onverterMap.put(double.class, DOUBLE_CONVERTER);
				onverterMap.put(String.class, STRING_CONVERTER);
			}
			Preconditions.checkArgument(
					onverterMap != null && onverterMap.size() > 0,
					"类型转换map不能为空");
			Preconditions.checkNotNull(workbook, "excel不能为空");
			return new SheetParser<T>(targetClass, onverterMap,
					workbook.getSheetAt(0));
		}
	}
}
