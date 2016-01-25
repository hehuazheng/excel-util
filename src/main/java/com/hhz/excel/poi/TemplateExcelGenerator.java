package com.hhz.excel.poi;

import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.common.collect.Maps;
import com.hhz.excel.annotation.SheetColumnAttribute;
import com.hhz.excel.poi.support.FieldNameIndexMapper;
import com.hhz.excel.poi.support.RowGenerator;
import com.hhz.excel.support.ExcelUtils;
import com.hhz.excel.support.FieldUtils;

public class TemplateExcelGenerator<T> extends ExcelGenerator<T> {
	private MyRowGenerator<T> myRowGenerator;

	public TemplateExcelGenerator(ExcelGeneratorFactory.Builder<T> builder) {
		super(builder.getWorkbook(), builder.getTargetClass());
	}

	@Override
	protected Sheet getProcessSheet() {
		return workbook.getSheetAt(0);
	}

	@Override
	RowGenerator<T> getRowGenerator() {
		if (myRowGenerator == null) {
			myRowGenerator = new MyRowGenerator<T>(this);
		}
		return myRowGenerator;
	}

	private static class MyRowGenerator<D> implements RowGenerator<D> {
		Map<Integer, FieldWrapper> indexFieldMap;
		private final Workbook workbook;
		private int columnCount = 0;
		private Map<String, CellStyle> cellStyleMap = Maps.newHashMap();

		public MyRowGenerator(TemplateExcelGenerator<D> generator) {
			this.workbook = generator.workbook;
			Row titleRow = ExcelUtils.getTitleRow(generator.getProcessSheet(),
					generator.targetClass);
			columnCount = titleRow.getPhysicalNumberOfCells();
			List<FieldWrapper> fieldWrapperList = FieldUtils
					.getFieldWrapperList(generator.targetClass);
			indexFieldMap = FieldNameIndexMapper.toIndexedMap(titleRow,
					fieldWrapperList);
		}

		private CellStyle createDateCellStyleIfNecessary(String dateFormat) {
			CellStyle cellStyle = cellStyleMap.get(dateFormat);
			if (cellStyle == null) {
				cellStyle = workbook.createCellStyle();
				CreationHelper createHelper = workbook.getCreationHelper();
				cellStyle.setDataFormat(createHelper.createDataFormat()
						.getFormat(dateFormat));
				cellStyleMap.put(dateFormat, cellStyle);
			}
			return cellStyle;
		}

		@Override
		public Row generate(Sheet sheet, D rowData) throws ExcelException {
			int currentRowNum = sheet.getPhysicalNumberOfRows();
			Row r = sheet.createRow(currentRowNum);
			for (int i = 0; i < columnCount; i++) {
				Cell cell = r.createCell(i);
				FieldWrapper fw = indexFieldMap.get(i);
				if (fw != null) {
					try {
						Object val = fw.getField().get(rowData);
						if (val == null) {
							continue;
						} else if (val instanceof Number) {
							cell.setCellValue(((Number) val).doubleValue());
						} else if (val instanceof Date) {
							SheetColumnAttribute sca = fw.getField()
									.getAnnotation(SheetColumnAttribute.class);
							if (sca != null) {
								String dateFormat = sca.dateFormat();
								cell.setCellStyle(createDateCellStyleIfNecessary(dateFormat));
							}
							cell.setCellValue((Date) val);
						} else if (val instanceof Boolean) {
							cell.setCellValue((Boolean) val);
						} else if (val instanceof String) {
							cell.setCellValue((String) val);
						} else {
							throw new ExcelException("不支持的类型" + val.toString());
						}
					} catch (ExcelException e) {
						throw e;
					} catch (Exception e) {
						throw new ExcelException("生成行失败", e);
					}
				}
			}
			return r;
		}
	}
}
