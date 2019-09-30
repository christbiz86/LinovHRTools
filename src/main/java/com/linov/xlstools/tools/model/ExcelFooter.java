package com.linov.xlstools.tools.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelFooter {

	private HSSFWorkbook workbook;
	private List<ExcelText> texts;

	protected ExcelFooter(HSSFWorkbook workbook) {
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

	public ExcelText removeText(int i) {
		return this.texts.remove(i);
	}

	public boolean removeText(ExcelText text) {
		return this.texts.remove(text);
	}

}
