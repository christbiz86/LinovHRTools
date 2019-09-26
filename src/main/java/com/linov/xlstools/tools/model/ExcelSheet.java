package com.linov.xlstools.tools.model;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelSheet {

	private ExcelHeader header;
	private ExcelTable table;
	private ExcelFooter footer;
	
	public ExcelSheet(HSSFWorkbook workbook, Sheet sheet) {
		this.header = new ExcelHeader(workbook);;
		this.table = new ExcelTable(workbook);
		this.footer = new ExcelFooter(workbook);
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
