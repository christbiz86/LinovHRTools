package com.linov.xlstools.tools.model;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelReport {

	private HSSFWorkbook workbook;
	private String fileName;
	private List<ExcelSheet> sheets;
	private Utils utils;
	
	public ExcelReport(String fileName) {
		this.workbook = new HSSFWorkbook();
		this.fileName = fileName;
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

	
}
