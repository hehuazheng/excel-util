package com.hhz.excel;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.hhz.excel.annotation.ExcelColumn;
import com.hhz.excel.annotation.ExcelModel;

public class ExcelReader {
	public static void main(String[] args) throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream(
				"d:/Book1.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
		AnnotationExcelDescriptor descriptor = new AnnotationExcelDescriptor(
				Cols.class);
		int rowCount = sheet.getPhysicalNumberOfRows();
		if (rowCount >= descriptor.getTitleRowIndex()) {
			descriptor
					.initFieldMap(sheet.getRow(descriptor.getTitleRowIndex()));
		}
		ColsExcel2007RowConverter converter = new ColsExcel2007RowConverter(
				descriptor);

		System.out.println("row count: " + rowCount);
		for (int i = descriptor.getTitleRowIndex() + 1; i <= rowCount; i++) {
			Row row = sheet.getRow(i);
			Cols cols = converter.convert(row);
			if (cols != null) {
				System.out.println(cols.getCol1() + " " + cols.getCol2() + " "
						+ cols.getCol3());
			}
		}
	}

	@ExcelModel
	static class Cols {
		@ExcelColumn("列1")
		private String col1;
		@ExcelColumn("列2")
		private String col2;
		@ExcelColumn("列3")
		private String col3;

		public String getCol1() {
			return col1;
		}

		public void setCol1(String col1) {
			this.col1 = col1;
		}

		public String getCol2() {
			return col2;
		}

		public void setCol2(String col2) {
			this.col2 = col2;
		}

		public String getCol3() {
			return col3;
		}

		public void setCol3(String col3) {
			this.col3 = col3;
		}
	}

	static class ColsExcel2007RowConverter implements
			Excel2007RowConverter<ExcelReader.Cols> {

		private Map<Integer, Field> fieldMap;

		ColsExcel2007RowConverter(AnnotationExcelDescriptor descriptor) {
			this.fieldMap = descriptor.getFieldMap();
		}

		@Override
		public Cols convert(Row source) {
			if (source != null) {
				Cols cols = new Cols();
				for (int i = 1; i <= source.getPhysicalNumberOfCells(); i++) {
					Field f = fieldMap.get(i);
					if (f != null) {
						Cell cell = source.getCell(i);
						f.setAccessible(true);
						try {
							f.set(cols, cell.getStringCellValue());
						} catch (IllegalArgumentException
								| IllegalAccessException e) {
							e.printStackTrace();
						}
					}
				}
				return cols;
			}
			return null;
		}

	}

	<T> T convert(Row row) {
		Excel2007RowConverter<Cols> converter = new Excel2007RowConverter<ExcelReader.Cols>() {
			@Override
			public Cols convert(Row row) {
				if (row != null) {
					int colCount = row.getPhysicalNumberOfCells();
					for (int j = 0; j < colCount; j++) {
						Cell cell = row.getCell(j);
						if (cell != null) {
							System.out.print(cell.getStringCellValue() + " ");
						} else {
							System.out.print("null ");
						}
					}
				}
				// TODO Auto-generated method stub
				return null;
			}
		};

		return convert(row);
	}

}
