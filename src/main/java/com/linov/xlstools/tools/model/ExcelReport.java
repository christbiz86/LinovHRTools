package com.linov.xlstools.tools.model;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelReport {

	private HSSFWorkbook workbook;
	private String name;
	private List<ExcelSheet> sheets;
	
	public ExcelReport(String name) {
		this.workbook = new HSSFWorkbook();
		this.sheets = new ArrayList<ExcelSheet>();
		this.name = name;
	}

	public void addSheet(String sheetName) {
		ExcelSheet sheet = new ExcelSheet(this.workbook, sheetName);
		this.sheets.add(sheet);
	}
	
	public void createFile() throws IOException {
		Utils.generateWorkbook(this.workbook, this.sheets);
		File currDir = new File(".");
		String path = currDir.getAbsolutePath();
		String fileLocation = path.substring(0, path.length() - 1) + this.name + ".xls";
		
		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		try {
			this.workbook.write(outputStream);
		} finally {
			outputStream.close();
		}
	}
	
	public void createFile(String path) throws IOException {
		Utils.generateWorkbook(this.workbook, this.sheets);
		String fileLocation = path + this.name + ".xls";

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		try {
			this.workbook.write(outputStream);
		} finally {
			outputStream.close();
		}
	}
	
	public byte[] toByteArray() throws IOException {
		Utils.generateWorkbook(this.workbook, this.sheets);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
		    this.workbook.write(bos);
		} finally {
		    bos.close();
		}
		byte[] bytes = bos.toByteArray();
		return bytes;
	}
	
	public void close() throws IOException {
		this.workbook.close();
	}
	
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public ExcelSheet getSheet(Integer index) {
		return this.sheets.get(index);
	}

	public List<ExcelSheet> getSheets() {
		return this.sheets;
	}
}
