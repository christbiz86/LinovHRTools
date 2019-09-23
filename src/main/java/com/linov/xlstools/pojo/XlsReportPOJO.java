package com.linov.xlstools.pojo;

import java.util.List;

public class XlsReportPOJO {

	private String fileName;
	private List<SheetPOJO> sheets;
	
	public String getFileName() {
		return fileName;
	}
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}
	public List<SheetPOJO> getSheets() {
		return sheets;
	}
	public void setSheets(List<SheetPOJO> sheetPOJOs) {
		this.sheets = sheetPOJOs;
	}
	
	
}
