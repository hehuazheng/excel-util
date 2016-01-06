package com.hhz.excel.poi.support;

import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

import com.hhz.excel.poi.CellConvertException;

public class DateCellToStringConverter implements CellConverter<String> {
	public static final String DEFAULT_DATE_FORMAT = "yyyy-mm-dd hh:MM:ss";

	private String dateFormat;

	public DateCellToStringConverter() {
		this(DEFAULT_DATE_FORMAT);
	}

	public DateCellToStringConverter(String dateFormat) {
		this.dateFormat = dateFormat;
	}

	@Override
	public String convert(Cell cell, boolean stopOnError)
			throws CellConvertException {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				return new SimpleDateFormat(dateFormat).format(cell
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
			if (!stopOnError) {
				throw new CellConvertException("不识别的excel cell类型");
			}
		}
		return null;
	}

}
