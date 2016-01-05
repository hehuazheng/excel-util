package com.hhz.common;

public interface Converter<S, T> {
	T convert(S source);
}
