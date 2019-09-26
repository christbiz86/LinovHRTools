package com.linov.xlstools.tools.model;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelField {
	
	private String name;
	private ExcelStyle style;
	private HSSFFont font;
	
	protected ExcelField (HSSFWorkbook workbook, String name) {
		this.name = name;
		this.style = new ExcelStyle(workbook);
		this.font = workbook.createFont();
	}
	
	public String getName() {
		return name;
	}
	public void setName(String Name) {
		this.name = Name;
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
