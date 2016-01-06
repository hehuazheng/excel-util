package com.hhz.excel;

import java.lang.reflect.Field;

public class FieldWrapper {
	public FieldWrapper(Field field, int index) {
		field.setAccessible(true);
		this.field = field;
		this.index = index;
	}

	private Field field;
	// 列标
	private int index;

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
}
