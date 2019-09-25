package com.linov.xlstools.tools.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelFooter {

	private HSSFWorkbook workbook;
	private List<ExcelText> texts;

	public ExcelFooter(HSSFWorkbook workbook) {
		this.workbook = workbook;
		texts = new ArrayList<ExcelText>();
	}

	public List<ExcelText> getTexts() {
		return texts;
	}

	public void setTexts(List<ExcelText> texts) {
		this.texts = texts;
	}

	public void addText(Object value) {
		ExcelText text = new ExcelText(this.workbook, value);
		this.texts.add(text);
	}
}
