package com.hhz.excel;

import java.util.Map;

import com.google.common.collect.Maps;

public abstract class AbstractExcelDescripor {
	private int titleRowIndex;
	protected Map<Integer, FieldWrapper> fieldMap;

	public AbstractExcelDescripor() {
		fieldMap = Maps.newHashMap();
	}

	public int getTitleRowIndex() {
		return titleRowIndex;
	}

	public void setTitleRowIndex(int titleRowIndex) {
		this.titleRowIndex = titleRowIndex;
	}

	public Map<Integer, FieldWrapper> getFieldMap() {
		return fieldMap;
	}

	public void setFieldMap(Map<Integer, FieldWrapper> fieldMap) {
		this.fieldMap = fieldMap;
	}

	enum ExtractType {
		byName, byIndex;
	}
}
