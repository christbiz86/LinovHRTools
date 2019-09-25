package com.linov.xlstools.tools.model;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelReport {

	private HSSFWorkbook workbook;
	private String fileName;
	private ExcelHeader header;
	private ExcelTable table;
	private ExcelFooter footer;
	private Utils utils;
	
	public ExcelReport(String fileName) {
		this.workbook = new HSSFWorkbook();
		this.fileName = fileName;
		this.header = new ExcelHeader(this.workbook);;
		this.table = new ExcelTable(this.workbook);
		this.footer = new ExcelFooter(this.workbook);
		this.utils = new Utils(this.workbook);
	}

	public void write() throws IOException {
		File currDir = new File(".");
		String path = currDir.getAbsolutePath();
		String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		this.workbook.write(outputStream);
	}
	
	public void close() throws IOException {
		this.workbook.close();
	}
	
	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
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
