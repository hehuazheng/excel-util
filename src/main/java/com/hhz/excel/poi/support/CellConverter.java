package com.hhz.excel.poi.support;

import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

import com.hhz.excel.poi.CellConvertException;

public interface CellConverter<S> {

	S convert(Cell cell) throws CellConvertException;

	public static CellConverter<String> CELL_TO_STRING_CONVERTER = new DateCellToStringConverter();

	public static CellConverter<Integer> CELL_TO_INTEGER_CONVERTER = new CellConverter<Integer>() {
		@Override
		public Integer convert(Cell cell) throws CellConvertException {
			int cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_FORMULA) {
				cellType = cell.getCachedFormulaResultType();
			}
			if (cellType == Cell.CELL_TYPE_NUMERIC
					&& !HSSFDateUtil.isCellDateFormatted(cell)) {
				return (int) cell.getNumericCellValue();
			}
			return null;
		}
	};

	public static CellConverter<Long> CELL_TO_LONG_CONVERTER = new CellConverter<Long>() {
		@Override
		public Long convert(Cell cell) throws CellConvertException {
			int cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_FORMULA) {
				cellType = cell.getCachedFormulaResultType();
			}
			if (cellType == Cell.CELL_TYPE_NUMERIC
					&& !HSSFDateUtil.isCellDateFormatted(cell)) {
				double value = cell.getNumericCellValue();

				if (Math.floor(value) == value) {
					return Math.round(value);
				}
				throw new CellConvertException(value + "不能转换为long");
			}
			return null;
		}
	};

	public static CellConverter<Double> CELL_TO_DOUBLE_CONVERTER = new CellConverter<Double>() {
		@Override
		public Double convert(Cell cell) throws CellConvertException {
			int cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_FORMULA) {
				cellType = cell.getCachedFormulaResultType();
			}
			if (cellType == Cell.CELL_TYPE_NUMERIC
					&& !HSSFDateUtil.isCellDateFormatted(cell)) {
				return cell.getNumericCellValue();
			}
			return null;
		}
	};

	public static CellConverter<Date> CELL_TO_DATE_CONVERTER = new CellConverter<Date>() {
		@Override
		public Date convert(Cell cell) throws CellConvertException {
			int cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_FORMULA) {
				cellType = cell.getCachedFormulaResultType();
			}
			if (cellType == Cell.CELL_TYPE_NUMERIC
					&& HSSFDateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			}
			return null;
		}
	};
}