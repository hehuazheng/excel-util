package com.hhz.excel.poi;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;
import com.hhz.excel.support.ExcelUtils;

public class ExcelGeneratorFactory {
	public static class Builder<T> {
		private Class<T> targetClass;
		private Workbook workbook;

		private Builder(Class<T> targetClass) {
			this.targetClass = targetClass;
		}

		public ExcelGenerator<T> build() {
			if (workbook == null) {
				workbook = new XSSFWorkbook();
				return new NewExcelGenerator<T>(this);
			} else {
				return new TemplateExcelGenerator<T>(this);
			}
		}

		public Builder<T> template(InputStream is) {
			Preconditions.checkArgument(workbook == null, "不允许设置重复的模板");
			workbook = ExcelUtils.getXSSFWorkbook(is);
			return this;
		}

		public Builder<T> template(String templateFile) {
			workbook = ExcelUtils.getXSSFWorkbook(templateFile);
			return this;
		}

		public Workbook getWorkbook() {
			return workbook;
		}

		public Class<T> getTargetClass() {
			return targetClass;
		}
	}

	public static <T> Builder<T> builder(Class<T> targetClass) {
		return new Builder<T>(targetClass);
	}

}
