package com.hhz.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface ExcelModel {
	/**
	 * excel 标题栏所在行数
	 */
	int titleRowIndex() default 0;

	RangeType rangeType() default RangeType.FIXED;

	public enum RangeType {
		FIXED, RELATIVE;
	}

	int minSheetRange() default 0;

	int maxSheetRange() default Integer.MAX_VALUE;
}
