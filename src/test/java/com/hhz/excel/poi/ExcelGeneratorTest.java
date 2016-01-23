package com.hhz.excel.poi;

import java.io.FileOutputStream;
import java.util.List;

import org.junit.Test;

import com.google.common.collect.Lists;
import com.hhz.excel.TestFileUtils;
import com.hhz.excel.annotation.SheetAttribute;
import com.hhz.excel.annotation.SheetColumnAttribute;

public class ExcelGeneratorTest {

	@SheetAttribute
	static class Data {
		@SheetColumnAttribute(title = "姓名")
		private String name;
		@SheetColumnAttribute(title = "地址")
		private String addr;

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getAddr() {
			return addr;
		}

		public void setAddr(String addr) {
			this.addr = addr;
		}
	}

	@Test
	public void testCreateTotallyNew() throws Exception {
		// .getResourceAsStream("te"));
		Data d1 = new Data();
		d1.setName("hzz");
		d1.setAddr("pj");
		Data d2 = new Data();
		d2.setName("name2");
		d2.setAddr("addr2");
		List<Data> list = Lists.newArrayList(d1, d2);
		ExcelGeneratorFactory
				.builder()
				.build()
				.process(list)
				.write(new FileOutputStream(TestFileUtils
						.getFilePath("testw-created.xlsx")));
	}

	@Test
	public void testCreateFromTemplate() throws Exception {
		Data d1 = new Data();
		d1.setName("hzz");
		d1.setAddr("pj");
		Data d2 = new Data();
		d2.setName("name2");
		d2.setAddr("addr2");
		List<Data> list = Lists.newArrayList(d1, d2);
		ExcelGeneratorFactory
				.builder()
				.template(TestFileUtils.getFilePath("testw.xlsx"))
				.build()
				.process(list)
				.write(new FileOutputStream(TestFileUtils
						.getFilePath("testw-created.xlsx")));
	}
}
