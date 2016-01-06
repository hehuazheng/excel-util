package com.hhz.excel.util;

import org.apache.poi.ss.usermodel.Cell;

public class CellUtil {
	public static String readCellValueToString(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case Cell.CELL_TYPE_BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		default:
			throw new RuntimeException("不识别的excel cell类型");
		}
	}
}
