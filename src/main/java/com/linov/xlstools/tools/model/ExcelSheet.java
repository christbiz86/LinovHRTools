package com.linov.xlstools.tools.model;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelSheet {

	private String name;
	private ExcelHeader header;
	private ExcelTable table;
	private ExcelFooter footer;
	
	public ExcelSheet(HSSFWorkbook workbook, String sheetName) {
		this.name = sheetName;
		this.header = new ExcelHeader(workbook);
		this.table = new ExcelTable(workbook);
		this.footer = new ExcelFooter(workbook);
	}

	public String getName() {
		return this.name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public ExcelHeader getHeader() {
		return header;
	}

	public ExcelTable getTable() {
		return table;
	}

	public ExcelFooter getFooter() {
		return footer;
	}
	
}
