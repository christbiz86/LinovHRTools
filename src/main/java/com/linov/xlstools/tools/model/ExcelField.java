package com.linov.xlstools.tools.model;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelField {
	
	private String name;
	private ExcelStyle style;
	private HSSFFont font;
	
	protected ExcelField (HSSFWorkbook workbook, String name) {
		this.name = name;
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
