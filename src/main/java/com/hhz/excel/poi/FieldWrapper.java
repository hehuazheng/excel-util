package com.hhz.excel.poi;

import java.lang.reflect.Field;

public class FieldWrapper {
	public FieldWrapper(Field field, String displayName, int index,
			boolean required) {
		field.setAccessible(true);
		this.field = field;
		this.displayName = displayName;
		this.required = required;
	}

	private Field field;
	private String displayName;
	private boolean required;

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

	public boolean isRequired() {
		return required;
	}

	public void setRequired(boolean required) {
		this.required = required;
	}
}
