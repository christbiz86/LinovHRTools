package com.linov.xlstools.tools.model;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelText {

	private Object value;
	private ExcelStyle style;
	private HSSFFont font;
	
	public ExcelText(HSSFWorkbook workbook, Object value) {
		this.value = value;
	}
	public Object getValue() {
		return value;
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

}
