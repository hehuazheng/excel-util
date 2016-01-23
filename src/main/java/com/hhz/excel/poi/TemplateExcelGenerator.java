package com.hhz.excel.poi;

import org.apache.poi.ss.usermodel.Sheet;

public class TemplateExcelGenerator extends ExcelGenerator {

	public TemplateExcelGenerator(ExcelGeneratorFactory.Builder builder) {
		super(builder.getWorkbook());
	}

	@Override
	protected Sheet getProcessSheet() {
		return workbook.getSheetAt(0);
	}

}
