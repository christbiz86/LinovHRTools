package com.linov.xlstools.tools.model;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelText {

	private Object value;
	private ExcelStyle style;
	private HSSFFont font;
	
	protected ExcelText(HSSFWorkbook workbook, Object value) {
		setTextValue(value);
		this.style = new ExcelStyle(workbook);
	}
	
	private void setTextValue(Object value) {
		if (value == null) {
			this.value = value;
		} else if (value.getClass() == LocalDateTime.class) {
			this.value = Date.from(((LocalDateTime) value).atZone(ZoneId.systemDefault()).toInstant());
		} else if (value.getClass() == LocalDate.class) {
			this.value = Date.from(((LocalDate) value).atStartOfDay(ZoneId.systemDefault()).toInstant());
		} else {
			this.value = value;
		}
	}

	public ValueType getValueType() {
		if (this.value == null) {
			return ValueType.NULL;
		} else if (this.value.getClass() == Integer.class || this.value.getClass() == Double.class
				|| this.value.getClass() == Float.class || this.value.getClass() == Long.class) {
			return ValueType.NUMERIC;
		} else if (this.value.getClass() == Date.class || this.value.getClass() == LocalDateTime.class
				|| this.value.getClass() == LocalDate.class) {
			return ValueType.DATE;
		} else if (this.value.getClass() == LocalTime.class) {
			return ValueType.TIME;
		} else if (this.value.getClass() == Boolean.class || this.value.getClass() == LocalTime.class) {
			return ValueType.BOOLEAN;
		} else {
			return ValueType.STRING;
		}
	}
	
	public Object getValue() {
		return this.value;
	}
	public void setValue(Object value) {
		this.value = value;
	}
	public ExcelStyle getStyle() {
		return style;
	}
	protected CellStyle getCellStyle() {
		return style.getCellStyle();
	}
	public void setStyle(ExcelStyle excelStyle) {
		this.style = excelStyle;
	}
	public HSSFFont getFont() {
		return font;
	}
	public void setFont(HSSFFont font) {
		this.font = font;
	}

	public enum ValueType {
		NULL,
		STRING,
		NUMERIC,
		DATE,
		TIME,
		BOOLEAN
	}
	public enum DateType {
		STRING,
		NUMERIC,
		DATE,
		BOOLEAN
	}
}
