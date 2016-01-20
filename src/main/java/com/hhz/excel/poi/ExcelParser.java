package com.hhz.excel.poi;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.Date;
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
import com.hhz.excel.poi.support.CellConverter;
import com.hhz.excel.support.AnnotationSheetDefinition;

public class ExcelParser<T> {
	private final AnnotationSheetDefinition descriptor;
	private final Class<T> targetClass;
	private Workbook workbook;

	private Map<Class<?>, CellConverter<?>> converterMap;

	private ExcelParser(Workbook workbook, Class<T> targetClass,
			Map<Class<?>, CellConverter<?>> converterMap, boolean stopOnError) {
		this.workbook = workbook;
		this.targetClass = targetClass;
		this.descriptor = new AnnotationSheetDefinition(targetClass);
		this.converterMap = converterMap;
	}

	public void initFieldMap(Row row) {
		if (descriptor.getFieldWrapperList().isEmpty()) {
			Preconditions.checkNotNull(row, "标题列不能为空");
			int cellCount = row.getPhysicalNumberOfCells();
			for (int i = 0; i < cellCount; i++) {
				Cell cell = row.getCell(i);
				if (cell != null) {
					String titleName = cell.getStringCellValue().trim();
					FieldWrapper fw = descriptor.getFieldByTitleName(titleName);
					if (fw != null) {
						fw.setIndex(i);
						descriptor.addFieldWrapper(fw);
					}
				}
			}
		}
	}

	public List<T> toList() throws ExcelException {
		List<T> list = Lists.newArrayList();
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			List<T> tmpList = Lists.newArrayList();
			Row titleRow = sheet.getRow(descriptor.getTitleRowIndex());
			initFieldMap(titleRow);
			int rowCount = sheet.getPhysicalNumberOfRows();
			for (int j = descriptor.getTitleRowIndex() + 1; j <= rowCount; j++) {
				Row row = sheet.getRow(j);
				try {
					T t = convert(row);
					if (t != null) {
						tmpList.add(t);
					}
				} catch (Exception e) {
					throw new ExcelException("解析excel异常", e);
				}
			}
			list.addAll(tmpList);
		}
		return list;
	}

	private T convert(Row source) throws Exception {
		if (source != null) {
			T obj = targetClass.newInstance();
			for (FieldWrapper fw : descriptor.getFieldWrapperList()) {
				Field f = fw.getField();
				Cell cell = source.getCell(fw.getIndex());
				try {
					setFieldValue(fw, obj, cell);
				} catch (Exception e) {
					e.printStackTrace();
					throw new ParseExcelException(f.getName() + "设置值时出错", e);
				}
			}
			return obj;
		}
		return null;
	}

	private void setFieldValue(FieldWrapper f, Object obj, Cell cell)
			throws Exception {
		Field field = f.getField();
		Class<?> clazz = field.getType();
		CellConverter<?> cellConverter = converterMap.get(clazz);
		if (cellConverter != null) {
			try {
				field.set(obj, cellConverter.convert(cell));
			} catch (Exception e) {
				if(f.isRequired()) {
					throw e;
				}
			}
		}
	}

	public static class ExcelParserBuilder<T> {
		private static final Logger LOGGER = LoggerFactory
				.getLogger(ExcelParserBuilder.class);

		private static final Map<Class<?>, CellConverter<?>> DEFAULT_CONVERTER_MAP;
		static {
			Map<Class<?>, CellConverter<?>> map = Maps.newHashMap();
			map.put(int.class, CellConverter.CELL_TO_INTEGER_CONVERTER);
			map.put(Integer.class, CellConverter.CELL_TO_INTEGER_CONVERTER);
			map.put(double.class, CellConverter.CELL_TO_DOUBLE_CONVERTER);
			map.put(Double.class, CellConverter.CELL_TO_DOUBLE_CONVERTER);
			map.put(String.class, CellConverter.CELL_TO_STRING_CONVERTER);
			map.put(Date.class, CellConverter.CELL_TO_DATE_CONVERTER);
			DEFAULT_CONVERTER_MAP = map;
		}

		private Class<T> targetClass;
		private boolean stopOnError = true;
		private Workbook workbook;

		private ExcelParserBuilder(Class<T> targetClass) {
			this.targetClass = targetClass;
		}

		public ExcelParserBuilder<T> setWorkbook(Workbook workbook) {
			this.workbook = workbook;
			return this;
		}

		public ExcelParserBuilder<T> setStopOnError(boolean stopOnError) {
			this.stopOnError = stopOnError;
			return this;
		}

		public ExcelParserBuilder<T> setFilePath(String filePath) {
			try {
				this.workbook = WorkbookFactory.create(new FileInputStream(
						filePath));
			} catch (Exception e) {
				LOGGER.error("生成workbook异常 " + filePath, e);
			}
			return this;
		}

		public static <T> ExcelParserBuilder<T> create(Class<T> targetClass) {
			return new ExcelParserBuilder<T>(targetClass);
		}

		public ExcelParser<T> build() {
			Preconditions.checkNotNull(targetClass, "targetClass不能为空");
			Preconditions.checkNotNull(workbook, "excel不能为空");
			return new ExcelParser<T>(workbook, targetClass,
					DEFAULT_CONVERTER_MAP, stopOnError);
		}
	}
}
