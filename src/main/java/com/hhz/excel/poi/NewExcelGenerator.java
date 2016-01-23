package com.hhz.excel.poi;

import org.apache.poi.ss.usermodel.Sheet;

public class NewExcelGenerator extends ExcelGenerator {
	public NewExcelGenerator(ExcelGeneratorFactory.Builder builder) {
		super(builder.getWorkbook());
	}

	@Override
	protected Sheet getProcessSheet() {
		if (workbook.getNumberOfSheets() == 0) {
			workbook.createSheet();
		}
		return workbook.getSheetAt(0);
	}
}
