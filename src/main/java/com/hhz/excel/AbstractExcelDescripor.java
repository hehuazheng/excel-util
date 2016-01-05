package com.hhz.excel;

import java.lang.reflect.Field;
import java.util.Map;

import com.google.common.collect.Maps;

public abstract class AbstractExcelDescripor {
	private int titleRowIndex;
	private ExtractType extractType;
	protected Map<Integer, Field> fieldMap;

	public AbstractExcelDescripor() {
		fieldMap = Maps.newHashMap();
	}

	public int getTitleRowIndex() {
		return titleRowIndex;
	}

	public void setTitleRowIndex(int titleRowIndex) {
		this.titleRowIndex = titleRowIndex;
	}

	public ExtractType getExtractType() {
		return extractType;
	}

	public void setExtractType(ExtractType extractType) {
		this.extractType = extractType;
	}

	public Map<Integer, Field> getFieldMap() {
		return fieldMap;
	}

	public void setFieldMap(Map<Integer, Field> fieldMap) {
		this.fieldMap = fieldMap;
	}

	enum ExtractType {
		byName, byIndex;
	}
}
