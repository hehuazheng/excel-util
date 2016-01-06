package com.hhz.excel;

import java.util.List;

import com.google.common.base.Preconditions;
import com.google.common.collect.Lists;

public abstract class AbstractSheetDefinition {
	private int titleRowIndex;
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

	public List<FieldWrapper> addFieldWrapper(FieldWrapper fieldWrapper) {
		conflictCheck(fieldWrapper);
		fieldWrapperList.add(fieldWrapper);
		return fieldWrapperList;
	}

	private void conflictCheck(FieldWrapper fieldWrapper) {
		for (FieldWrapper fw : fieldWrapperList) {
			Preconditions.checkArgument(
					fw.getField() != fieldWrapper.getField(), fw.getField()
							.getName() + "冲突");
			Preconditions.checkArgument(
					fw.getIndex() != fieldWrapper.getIndex(),
					"列标冲突" + fw.getIndex());
		}
	}

	public List<FieldWrapper> getFieldWrapperList() {
		return fieldWrapperList;
	}

	public void setFieldWrapperList(List<FieldWrapper> fieldWrapperList) {
		this.fieldWrapperList = fieldWrapperList;
	}
}
