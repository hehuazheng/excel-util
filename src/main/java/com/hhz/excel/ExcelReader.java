package com.hhz.excel;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.alibaba.fastjson.JSON;
import com.hhz.excel.ExcelParser.ExcelParserBuilder;
import com.hhz.excel.annotation.ExcelColumn;
import com.hhz.excel.annotation.ExcelModel;

public class ExcelReader {
	public static void main(String[] args) throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream(
				"d:/Book1.xlsx"));
		List<Cols> list = ExcelParserBuilder.create(Cols.class).build()
				.parse(wb);
		System.out.println(JSON.toJSONString(list));
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
}
