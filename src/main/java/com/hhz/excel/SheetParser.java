package com.hhz.excel;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
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
import com.google.common.primitives.Primitives;
import com.hhz.common.Converter;

public class SheetParser<T> {
	private final AnnotationExcelDescriptor descriptor;
	private final Class<T> targetClass;
	private Map<Class<?>, Converter<String, ?>> converterMap;
	private Sheet sheet;
	private static final String DEFAULT_DATE_FORMAT = "yyyy-mm-dd hh:MM:ss";

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
				for (FieldWrapper fw : descriptor.getFieldWrapperList()) {
					Field f = fw.getField();
					Cell cell = source.getCell(fw.getIndex());
					try {
						perfectMatch(f, r, cell);
					} catch (Exception e) {
						e.printStackTrace();
						throw new ParseExcelException(f.getName() + "设置值时出错", e);
					}
				}
				return r;
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return null;
	}

	private void perfectMatch(Field f, Object obj, Cell cell) throws Exception {
		Class<?> clazz = f.getType();
		System.out.println(clazz);
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				if (clazz == Date.class) {
					f.set(obj, cell.getDateCellValue());
				} else {
					vset(f, obj,
							new SimpleDateFormat(DEFAULT_DATE_FORMAT)
									.format(cell.getDateCellValue()));
				}
			} else {
				if (clazz == Double.class || clazz == double.class) {
					f.set(obj, cell.getNumericCellValue());
				} else if (clazz == int.class || clazz == Integer.class) {
					f.set(obj, (int) cell.getNumericCellValue());
				} else if (clazz == String.class) {
					f.set(obj, String.valueOf(cell.getNumericCellValue()));
				} else {
					vset(f, obj, String.valueOf(cell.getNumericCellValue()));
				}
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			if (clazz == boolean.class || clazz == Boolean.class) {
				f.set(obj, cell.getBooleanCellValue());
			} else {
				vset(f, obj, String.valueOf(cell.getBooleanCellValue()));
			}
			break;
		case Cell.CELL_TYPE_STRING:
			if (clazz == String.class) {
				f.set(obj, cell.getStringCellValue());
			} else {
				vset(f, obj, cell.getStringCellValue());
			}
			break;
		case Cell.CELL_TYPE_BLANK:
			System.out.println("空列 ");
			break;
		case Cell.CELL_TYPE_FORMULA:
			break;
		case Cell.CELL_TYPE_ERROR:
			break;
		default:
			throw new RuntimeException("不识别的excel cell类型");
		}
	}

	void vset(Field f, Object obj, String value)
			throws IllegalArgumentException, IllegalAccessException {
		Converter<String, ?> converter = this.converterMap.get(f.getType());
		f.set(obj, converter.convert(value));
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
