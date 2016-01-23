package com.hhz.excel.poi;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.Preconditions;
import com.hhz.excel.support.ExcelUtils;

public class ExcelGeneratorFactory {
	public static class Builder {
		private Workbook workbook;

		private Builder() {
		}

		public ExcelGenerator build() {
			if (workbook == null) {
				workbook = new XSSFWorkbook();
				return new NewExcelGenerator(this);
			} else {
				return new TemplateExcelGenerator(this);
			}
		}

		public Builder template(InputStream is) {
			Preconditions.checkArgument(workbook == null, "不允许设置重复的模板");
			workbook = ExcelUtils.getWorkbook(is);
			return this;
		}

		public Builder template(String templateFile) {
			workbook = ExcelUtils.getWorkbook(templateFile);
			return this;
		}

		public Workbook getWorkbook() {
			return workbook;
		}
	}

	public static Builder builder() {
		return new Builder();
	}

}
