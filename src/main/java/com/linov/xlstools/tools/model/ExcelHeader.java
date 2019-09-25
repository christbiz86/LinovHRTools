package com.linov.xlstools.tools.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelHeader {

	private HSSFWorkbook workbook;
	private List<ExcelText> texts;

	public ExcelHeader(HSSFWorkbook workbook) {
		this.workbook = workbook;
		this.texts = new ArrayList<ExcelText>();
	}

	public void addText(Object value) {
		ExcelText text = new ExcelText(this.workbook, value);
		this.texts.add(text);
	}
	
	public ExcelText getText(Integer index) {
		return texts.get(index);
	}
	
	public List<ExcelText> getTexts() {
		return texts;
	}

	public void setTexts(List<ExcelText> texts) {
		this.texts = texts;
	}

}
