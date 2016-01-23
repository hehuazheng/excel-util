package com.hhz.excel.support;

import java.util.List;

import com.google.common.collect.Lists;
import com.hhz.excel.poi.FieldWrapper;

public abstract class AbstractSheetDefinition {
	private int titleRowIndex = 1;
	protected List<FieldWrapper> fieldWrapperList;

	public AbstractSheetDefinition() {
		fieldWrapperList = Lists.newArrayList();
	}

	public int getTitleRowIndex() {
		return titleRowIndex;
	}

	public void setTitleRowIndex(int titleRowIndex) {
		this.titleRowIndex = titleRowIndex;
	}

	public List<FieldWrapper> getFieldWrapperList() {
		return fieldWrapperList;
	}
}
