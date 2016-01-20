package com.hhz.excel.poi;

import java.lang.reflect.Field;

public class FieldWrapper {
	public FieldWrapper(Field field, int index, boolean required) {
		field.setAccessible(true);
		this.field = field;
		this.index = index;
		this.required = required;
	}

	private Field field;
	// 列标
	private int index;
	private boolean required;

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public boolean isRequired() {
		return required;
	}

	public void setRequired(boolean required) {
		this.required = required;
	}
}
