package com.hhz.excel.poi;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

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
import com.hhz.excel.support.AnnotationSheetDefinition;

public class ExcelParser<T> {
	private static final Logger LOGGER = LoggerFactory
			.getLogger(ExcelParser.class);
	private final AnnotationSheetDefinition descriptor;
	private final Class<T> targetClass;
	private Workbook workbook;
	private static final String DEFAULT_DATE_FORMAT = "yyyy-mm-dd hh:MM:ss";
	private boolean ignoreFieldError = false;

	private ExcelParser(Class<T> targetClass, Workbook workbook) {
		this.targetClass = targetClass;
		this.descriptor = new AnnotationSheetDefinition(targetClass);
		this.workbook = workbook;
	}

	public void initFieldMap(Row row) {
		if (descriptor.getFieldWrapperList().isEmpty()) {
			Preconditions.checkNotNull(row, "标题列不能为空");
			int cellCount = row.getPhysicalNumberOfCells();
			for (int i = 1; i <= cellCount; i++) {
				Cell cell = row.getCell(i);
				if (cell != null) {
					String titleName = cell.getStringCellValue().trim();
					Field f = descriptor.getFieldByTitleName(titleName);
					if (f != null) {
						descriptor.addFieldWrapper(new FieldWrapper(f, i));
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
					if (ignoreFieldError) {
						LOGGER.warn("excel转换数据失败", e);
					} else {
						throw new ExcelException("解析excel异常", e);
					}
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
					setFieldValue(f, obj, cell);
				} catch (Exception e) {
					e.printStackTrace();
					throw new ParseExcelException(f.getName() + "设置值时出错", e);
				}
			}
			return obj;
		}
		return null;
	}

	private String toStringValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				return new SimpleDateFormat(DEFAULT_DATE_FORMAT).format(cell
						.getDateCellValue());
			} else {
				return String.valueOf(cell.getNumericCellValue());
			}
		case Cell.CELL_TYPE_BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_BLANK:
		case Cell.CELL_TYPE_FORMULA:
		case Cell.CELL_TYPE_ERROR:
		default:
			if (!ignoreFieldError) {
				throw new RuntimeException("不识别的excel cell类型");
			}
		}
		return null;
	}

	private Integer toIntegerValue(Cell cell) throws ExcelException {
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
				&& !HSSFDateUtil.isCellDateFormatted(cell)) {
			return (int) cell.getNumericCellValue();
		}
		if (!ignoreFieldError) {
			throw new ExcelException();
		}
		return null;
	}

	private Double toDoubleValue(Cell cell) throws ExcelException {
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
				&& !HSSFDateUtil.isCellDateFormatted(cell)) {
			return cell.getNumericCellValue();
		}
		if (!ignoreFieldError) {
			throw new ExcelException();
		}
		return null;
	}

	private Date toDateValue(Cell cell) throws ExcelException {
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
				&& HSSFDateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue();
		}
		if (!ignoreFieldError) {
			throw new ExcelException();
		}
		return null;
	}

	private void setFieldValue(Field f, Object obj, Cell cell) throws Exception {
		Class<?> clazz = f.getType();
		if (clazz == String.class) {
			f.set(obj, toStringValue(cell));
		} else if (clazz == int.class || clazz == Integer.class) {
			f.set(obj, toIntegerValue(cell));
		} else if (clazz == double.class || clazz == Double.class) {
			f.set(obj, toDoubleValue(cell));
		} else if (clazz == Date.class) {
			f.set(obj, toDateValue(cell));
		}
	}

	public static class SheetParserBuilder<T> {
		private static final Logger LOGGER = LoggerFactory
				.getLogger(SheetParserBuilder.class);
		private Class<T> targetClass;
		private Workbook workbook;

		private SheetParserBuilder(Class<T> targetClass) {
			this.targetClass = targetClass;
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

		public ExcelParser<T> build() {
			Preconditions.checkNotNull(targetClass, "targetClass不能为空");
			Preconditions.checkNotNull(workbook, "excel不能为空");
			return new ExcelParser<T>(targetClass, workbook);
		}
	}
}
