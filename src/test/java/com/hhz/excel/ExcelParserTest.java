package com.hhz.excel;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import com.alibaba.fastjson.JSON;
import com.hhz.excel.ExcelParser.SheetParserBuilder;
import com.hhz.excel.annotation.SheetColumn;
import com.hhz.excel.annotation.SheetDescription;

public class ExcelParserTest {
	@SheetDescription
	static class Cols {
		@SheetColumn("列1")
		private String col1;
		@SheetColumn("列2")
		private String col2;
		@SheetColumn("列3")
		private String col3;
		@SheetColumn("列4")
		private double col4;
		@SheetColumn("列5")
		private String col5;
		@SheetColumn("列6")
		private int col6;
		@SheetColumn("列7")
		private Integer col7;
		@SheetColumn("列8")
		private Integer col8;

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

		public String getCol5() {
			return col5;
		}

		public void setCol5(String col5) {
			this.col5 = col5;
		}

		public int getCol6() {
			return col6;
		}

		public void setCol6(int col6) {
			this.col6 = col6;
		}

		public Integer getCol7() {
			return col7;
		}

		public void setCol7(Integer col7) {
			this.col7 = col7;
		}

		public Integer getCol8() {
			return col8;
		}

		public void setCol8(Integer col8) {
			this.col8 = col8;
		}
	}

	@Test
	public void testParse() throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream(TestFileUtils
				.getFilePath("test.xlsx")));
		List<Cols> list = SheetParserBuilder.create(Cols.class).setWorkbook(wb)
				.build().toList();
		System.out.println(JSON.toJSONString(list));
	}

}
