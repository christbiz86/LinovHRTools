package com.linov.xlstools.tools.model;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelRecord {

	private HSSFWorkbook workbook;
	private Map<String, ExcelText> record;

	protected ExcelRecord(HSSFWorkbook workbook, List<String> keys, List<Object> values) {
		this.workbook = workbook;
		this.record = new LinkedHashMap<String, ExcelText>();
		Integer i = 0;
		for (String key : keys) {
			ExcelText text = new ExcelText(this.workbook, values.get(i));
			this.record.put(key, text);
			i++;
		}
	}

	public Map<String, ExcelText> get() {
		return record;
	}
	
	public Integer size() {
		return record.size();
	}
	
	public ExcelText getText(String key) {
		return this.record.get(key);
	}
	
	public void set(String key, ExcelText text) {
		this.record.put(key, text);
	}

	public void setValue(String key, Object value) {
		ExcelText text = this.getText(key);
		text.setValue(value);
		this.record.put(key, text);
	}

}
