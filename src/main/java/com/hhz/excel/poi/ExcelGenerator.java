package com.hhz.excel.poi;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.hhz.excel.poi.support.RowGenerator;
import com.hhz.excel.poi.support.RowGenerators;

public abstract class ExcelGenerator {
	protected Workbook workbook;

	protected abstract Sheet getProcessSheet();

	public ExcelGenerator(Workbook workbook) {
		this.workbook = workbook;
	}

	public <D> Workbook process(List<D> list) throws ExcelException {
		return process(list, RowGenerators.getDefaultGenerator());
	}

	public <D> Workbook process(List<D> list, RowGenerator rg)
			throws ExcelException {
		for (D d : list) {
			rg.generate(getProcessSheet(), d);
		}
		return workbook;
	}
}
