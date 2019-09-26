package com.linov.xlstools.tools.model;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelRecord {

	private HSSFWorkbook workbook;
	private Map<String, ExcelText> values;

	protected ExcelRecord(HSSFWorkbook workbook, List<String> keys, List<Object> values) {
		this.workbook = workbook;
		this.values = new HashMap<String, ExcelText>();
		Integer i = 0;
		for (String key : keys) {
			ExcelText text = new ExcelText(this.workbook, values.get(i));
			this.values.put(key, text);
		}
	}

	public Map<String, ExcelText> get() {
		return values;
	}
	
	public Integer size() {
		return values.size();
	}
	
	public ExcelText getValueOf(String key) {
		return this.values.get(key);
	}
	
	public void put(String key, ExcelText text) {
		this.values.put(key, text);
	}

}
