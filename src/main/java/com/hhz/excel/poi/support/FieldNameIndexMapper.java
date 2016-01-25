package com.hhz.excel.poi.support;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.google.common.base.Preconditions;
import com.google.common.collect.Maps;
import com.hhz.excel.poi.FieldWrapper;

public class FieldNameIndexMapper {
	public static Map<Integer, FieldWrapper> toIndexedMap(Row row,
			List<FieldWrapper> fieldWrapperList) {
		Preconditions.checkArgument(fieldWrapperList != null
				&& fieldWrapperList.size() > 0, "fieldWrapperList不能为空");
		Map<Integer, FieldWrapper> fieldIndexMap = Maps.newHashMap();
		Map<String, FieldWrapper> titleNameKeydMap = toTitleKeydMap(fieldWrapperList);
		int cellCount = row.getPhysicalNumberOfCells();
		for (int i = 0; i < cellCount; i++) {
			Cell cell = row.getCell(i);
			if (cell != null) {
				String titleName = cell.getStringCellValue().trim();
				FieldWrapper fw = titleNameKeydMap.get(titleName);
				if (fw != null) {
					fieldIndexMap.put(i, fw);
				}
			}
		}
		return fieldIndexMap;
	}

	private static Map<String, FieldWrapper> toTitleKeydMap(
			List<FieldWrapper> fieldWrapperList) {
		Map<String, FieldWrapper> map = Maps.newHashMap();
		for (FieldWrapper fw : fieldWrapperList) {
			map.put(fw.getDisplayName(), fw);
		}
		return map;
	}
}
