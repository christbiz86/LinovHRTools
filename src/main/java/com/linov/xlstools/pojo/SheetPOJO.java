package com.linov.xlstools.pojo;

import java.util.List;

public class SheetPOJO {

	private String sheetName;
	private List<String> rptHeader;
	private List<String> gridHeader;
	private List<String[]> content;
	private List<String> footer;
	
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public List<String> getRptHeader() {
		return rptHeader;
	}
	public void setRptHeader(List<String> rptHeader) {
		this.rptHeader = rptHeader;
	}
	public List<String> getGridHeader() {
		return gridHeader;
	}
	public void setGridHeader(List<String> gridHeader) {
		this.gridHeader = gridHeader;
	}
	public List<String[]> getContent() {
		return content;
	}
	public void setContent(List<String[]> content) {
		this.content = content;
	}
	public List<String> getFooter() {
		return footer;
	}
	public void setFooter(List<String> footer) {
		this.footer = footer;
	}
	
}
