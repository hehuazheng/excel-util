package com.hhz.excel.poi.support;

import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

import com.hhz.excel.poi.CellConvertException;

public class CellToStringConverter implements CellConverter<String> {
	public static final String DEFAULT_DATE_FORMAT = "yyyy-mm-dd hh:MM:ss";

	@Override
	public String convert(Cell cell) throws CellConvertException {
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					return new SimpleDateFormat(DEFAULT_DATE_FORMAT)
							.format(cell.getDateCellValue());
				} else {
					Double doubleValue = cell.getNumericCellValue();
					if (doubleValue == doubleValue.intValue()) {
						return String.valueOf(doubleValue.intValue());
					} else {
						return String.valueOf(cell.getNumericCellValue());
					}
				}
			case Cell.CELL_TYPE_BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			case Cell.CELL_TYPE_BLANK:
			case Cell.CELL_TYPE_FORMULA:
			case Cell.CELL_TYPE_ERROR:
			default:
			}
		}
		return null;
	}

}
