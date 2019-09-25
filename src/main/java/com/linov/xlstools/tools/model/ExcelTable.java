package com.linov.xlstools.tools.model;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelTable {

	private HSSFWorkbook workbook;
	private List<ExcelField> fields;
	private List<ExcelRecord> records;
	
	public ExcelTable(HSSFWorkbook workbook) {
		this.workbook = workbook;
		this.fields = new ArrayList<ExcelField>();
		this.records = new ArrayList<ExcelRecord>();
	}

	public List<ExcelField> getFields() {
		return fields;
	}

	public void addField(String fieldName) {
		if (!this.fields.isEmpty()) {
			List<String> keys = new ArrayList<String>();
			for (ExcelField field : fields) {
				keys.add(field.getName());
			}
			if (keys.contains(fieldName)) {
				throw new IllegalArgumentException("Field " + fieldName + "is already exist");
			}
		}
		ExcelField field = new ExcelField(this.workbook, fieldName);
		this.fields.add(field);
	}

	public List<ExcelRecord> getRecords() {
		return records;
	}

	public void addRecord(List<Object> values) {
		if (this.fields.size() != values.size()) {
			throw new IllegalArgumentException("Fields and values size is not match");
		} else if (this.fields.isEmpty()) {
			throw new IllegalArgumentException("No field exists in this table");
		} 
		List<String> keys = new ArrayList<String>();
		for (ExcelField field : this.fields) {
			keys.add(field.getName());
		}
		ExcelRecord record = new ExcelRecord(this.workbook, keys, values);
		records.add(record);
	}
}
