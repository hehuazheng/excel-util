package com.hhz.excel;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import com.alibaba.fastjson.JSON;
import com.hhz.excel.ExcelParser.ExcelParserBuilder;
import com.hhz.excel.annotation.SheetColumn;
import com.hhz.excel.annotation.SheetModel;

public class ExcelParserTest {
	@SheetModel
	static class Cols {
		@SheetColumn("列1")
		private String col1;
		@SheetColumn("列2")
		private String col2;
		@SheetColumn("列3")
		private String col3;
		@SheetColumn("列4")
		private double col4;

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

		public double getCol4() {
			return col4;
		}

		public void setCol4(double col4) {
			this.col4 = col4;
		}
	}

	@Test
	public void testParse() throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream(TestFileUtils
				.getFilePath("test.xlsx")));
		List<Cols> list = ExcelParserBuilder.create(Cols.class).build()
				.parse(wb);
		System.out.println(JSON.toJSONString(list));
	}

}
