package com.hhz.excel.poi;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.hhz.excel.poi.support.CellConverter;

public abstract class ExcelParser<T> {
	private static final Logger LOGGER = LoggerFactory
			.getLogger(ExcelParser.class);

	private static final Map<Class<?>, CellConverter<?>> DEFAULT_CONVERTER_MAP;
	static {
		Map<Class<?>, CellConverter<?>> map = Maps.newHashMap();
		map.put(int.class, CellConverter.CELL_TO_INTEGER_CONVERTER);
		map.put(Integer.class, CellConverter.CELL_TO_INTEGER_CONVERTER);
		map.put(long.class, CellConverter.CELL_TO_LONG_CONVERTER);
		map.put(Long.class, CellConverter.CELL_TO_LONG_CONVERTER);
		map.put(double.class, CellConverter.CELL_TO_DOUBLE_CONVERTER);
		map.put(Double.class, CellConverter.CELL_TO_DOUBLE_CONVERTER);
		map.put(String.class, CellConverter.CELL_TO_STRING_CONVERTER);
		map.put(Date.class, CellConverter.CELL_TO_DATE_CONVERTER);
		DEFAULT_CONVERTER_MAP = map;
	}

	protected Workbook workbook;
	protected final Class<T> targetClass;
	private boolean multiSheet = false;
	private Map<Class<?>, CellConverter<?>> converterMap = DEFAULT_CONVERTER_MAP;

	public ExcelParser(Workbook workbook, Class<T> targetClass,
			boolean multiSheetEnabled) {
		this.workbook = workbook;
		this.targetClass = targetClass;
		this.multiSheet = multiSheetEnabled;
	}

	protected void setFieldValue(FieldWrapper f, Object obj, Cell cell)
			throws Exception {
		if (cell != null) {
			Field field = f.getField();
			Class<?> clazz = field.getType();
			CellConverter<?> cellConverter = converterMap.get(clazz);
			if (cellConverter != null) {
				try {
					field.set(obj, cellConverter.convert(cell));
				} catch (Exception e) {
					if (f.isRequired()) {
						LOGGER.error(field.getName() + "为空", e);
						throw e;
					}
					LOGGER.warn(field.getName() + "为空", e);
				}
			}
		}
	}

	protected T newRowModel() throws ExcelException {
		try {
			return targetClass.newInstance();
		} catch (Exception e) {
			e.printStackTrace();
			throw new ExcelException("生成列对象失败", e);
		}
	}

	protected abstract T convert(Row source) throws ExcelException;

	protected abstract List<T> processOneSheet(Sheet sheet)
			throws ExcelException;

	public List<T> toList() throws ExcelException {
		if (!multiSheet) {
			return processOneSheet(workbook.getSheetAt(0));
		} else {
			List<T> list = Lists.newArrayList();
			int sheetCount = workbook.getNumberOfSheets();
			for (int i = 0; i < sheetCount; i++) {
				Sheet sheet = workbook.getSheetAt(i);
				list.addAll(processOneSheet(sheet));
			}
			return list;
		}
	}
}
