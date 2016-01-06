package com.hhz.excel;

import java.lang.reflect.Field;

public class FieldWrapper {
	public FieldWrapper(Field field, String displayName) {
		this.field = field;
		this.displayName = displayName;
	}

	private Field field;
	private String displayName;

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

	public String getDisplayName() {
		return displayName;
	}

	public void setDisplayName(String displayName) {
		this.displayName = displayName;
	}
}
