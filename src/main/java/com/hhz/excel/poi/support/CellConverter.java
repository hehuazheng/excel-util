package com.hhz.excel.poi.support;

import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

import com.hhz.excel.poi.CellConvertException;

public interface CellConverter<S> {

	S convert(Cell cell, boolean stopOnError) throws CellConvertException;

	public static CellConverter<String> CELL_TO_STRING_CONVERTER = new DateCellToStringConverter();

	public static CellConverter<Integer> CELL_TO_INTEGER_CONVERTER = new CellConverter<Integer>() {
		@Override
		public Integer convert(Cell cell, boolean stopOnError)
				throws CellConvertException {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
					&& !HSSFDateUtil.isCellDateFormatted(cell)) {
				return (int) cell.getNumericCellValue();
			}
			if (!stopOnError) {
				throw new CellConvertException();
			}
			return null;
		}
	};

	public static CellConverter<Double> CELL_TO_DOUBLE_CONVERTER = new CellConverter<Double>() {
		@Override
		public Double convert(Cell cell, boolean stopOnError)
				throws CellConvertException {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
					&& !HSSFDateUtil.isCellDateFormatted(cell)) {
				return cell.getNumericCellValue();
			}
			if (!stopOnError) {
				throw new CellConvertException();
			}
			return null;
		}
	};

	public static CellConverter<Date> CELL_TO_DATE_CONVERTER = new CellConverter<Date>() {
		@Override
		public Date convert(Cell cell, boolean stopOnError)
				throws CellConvertException {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
					&& HSSFDateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			}
			if (!stopOnError) {
				throw new CellConvertException();
			}
			return null;
		}
	};
}