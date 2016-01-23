package com.hhz.excel.poi;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.base.Preconditions;
import com.hhz.excel.support.ExcelUtils;

public class ExcelParserFactory {
	public static class Builder<T> {
		private Workbook workbook;
		private Class<T> targetClass;
		private boolean multipleSheetEnabled = false;

		public Builder<T> workbook(String filePath) {
			Preconditions.checkArgument(workbook == null, "不允许重复设置workbook");
			this.workbook = ExcelUtils.getXSSFWorkbook(filePath);
			return this;
		}

		public Builder<T> workbook(InputStream inputStream) {
			Preconditions.checkArgument(workbook == null, "不允许重复设置workbook");
			this.workbook = ExcelUtils.getXSSFWorkbook(inputStream);
			return this;
		}

		public boolean isMultipleSheetEnabled() {
			return multipleSheetEnabled;
		}

		public void multipleSheetEnabled(boolean multipleSheetEnabled) {
			this.multipleSheetEnabled = multipleSheetEnabled;
		}

		public ExcelParser<T> build() {
			Preconditions.checkNotNull(targetClass, "targetClass不能为空");
			Preconditions.checkNotNull(workbook, "excel不能为空");
			return new AnnotationExcelParser<T>(this);
		}

		public Workbook getWorkbook() {
			return workbook;
		}

		public Class<T> getTargetClass() {
			return targetClass;
		}

		private Builder(Class<T> targetClass) {
			this.targetClass = targetClass;
		}
	}

	public static <T> Builder<T> builder(Class<T> targetClass) {
		return new Builder<T>(targetClass);
	}
}
