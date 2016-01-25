package com.hhz.excel.poi;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.hhz.excel.poi.support.RowGenerator;

public abstract class ExcelGenerator<T> {
	protected final Class<T> targetClass;
	protected Workbook workbook;

	protected abstract Sheet getProcessSheet();

	public ExcelGenerator(Workbook workbook, Class<T> targetClass) {
		this.workbook = workbook;
		this.targetClass = targetClass;
	}

	abstract RowGenerator<T> getRowGenerator();

	public Workbook addRows(List<T> list) throws ExcelException {
		RowGenerator<T> rg = getRowGenerator();
		for (T d : list) {
			rg.generate(getProcessSheet(), d);
		}
		return workbook;
	}
}
