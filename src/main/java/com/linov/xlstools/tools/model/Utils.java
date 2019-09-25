package com.linov.xlstools.tools.model;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Utils {

	private Integer rowIndex;
	private HSSFWorkbook workbook;
	private Sheet sheet;
	
	public Utils(HSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public void createReportHeader(ExcelHeader header) {
		
		for (int i = 0; i < header.getTexts().size(); i++) {
			Row row = this.sheet.createRow(this.rowIndex);

			Cell headerCell = row.createCell(0);
			headerCell.setCellValue(header.getText(i).getValue());
			headerCell.setCellStyle(header.getText(i).getCellStyle());
			this.rowIndex++;
		}
	}

	public void createGridHeader(List<String> gridHeader) {
		Row row = this.sheet.createRow(this.rowIndex);

		for (Integer i = 0; i < gridHeader.size(); i++) {
			CellStyle style = setFieldBorder(i, gridHeader.size());
			
			Cell cell = row.createCell(i);
			cell.setCellValue(gridHeader.get(i));
			cell.setCellStyle(style);
		}
		this.rowIndex++;
	}

	public CellStyle setFieldBorder(Integer currentColumn, Integer maxColumn) {
		CellStyle style = this.workbook.createCellStyle();

		if (isLeftMost(currentColumn)) {
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
		} else if (isRightMost(currentColumn, maxColumn)) {
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
		} else {
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);
		}
		
		return style;
	}

	public boolean isRightMost(Integer currentColumn, Integer maxColumn) {
		return currentColumn == maxColumn - 1;
	}

	public boolean isLeftMost(Integer currentColumn) {
		return currentColumn == 0;
	}

	public void createContent(List<String[]> content) {
		
		for (Integer i = 0; i < content.size(); i++) {
			Row row = this.sheet.createRow(this.rowIndex);
			Integer j = 0;
			
			for (String value : content.get(i)) {
				CellStyle style = stylizeContent(i, j, content.size(), content.get(i).length);
				Cell cell = row.createCell(j);
				cell.setCellValue(value);
				cell.setCellStyle(style);
				j++;
			}
			
			this.rowIndex++;
		}
	}

	public CellStyle stylizeContent(Integer currentRow, Integer currentColumn, Integer maxRow, Integer maxColumn) {
		CellStyle style = this.workbook.createCellStyle();

		if (isTopMost(currentRow)) {
			if (isLeftMost(currentColumn)) {
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
			} else if (currentColumn < maxColumn - 1) {
				style.setBorderTop(BorderStyle.THIN);
			} else if (isRightMost(currentColumn, maxColumn)) {
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			}
		} else if (isBottomMost(currentRow, maxRow)) { 
			if (isLeftMost(currentColumn)) {
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
			} else if (currentColumn < maxColumn - 1) {
				style.setBorderBottom(BorderStyle.THIN);
			} else if (isRightMost(currentColumn, maxColumn)) {
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			}
		} else {
			if (isLeftMost(currentColumn)) {
				style.setBorderLeft(BorderStyle.THIN);
			} else if (isRightMost(currentColumn, maxColumn)) {
				style.setBorderRight(BorderStyle.THIN);
			}
		}

		HSSFFont font = ((HSSFWorkbook) this.workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 12);
		style.setFont(font);
		
		style.setWrapText(true);
		return style;
	}

	public boolean isBottomMost(Integer currentRow, Integer maxRow) {
		return currentRow == maxRow - 1;
	}

	public boolean isTopMost(Integer currentRow) {
		return currentRow == 0;
	}

	public void createFooter(List<String> footer) {
		CellStyle footerStyle = stylizeRptHeader();
		
		for (int i = 0; i < footer.size(); i++) {
			Row row = this.sheet.createRow(this.rowIndex);

			Cell footerCell = row.createCell(0);
			footerCell.setCellValue(footer.get(i));
			footerCell.setCellStyle(footerStyle);
			this.rowIndex++;
		}
		this.rowIndex++;
	}
	
	public void createFile(String fileName) throws FileNotFoundException, IOException {
		File currDir = new File(".");
		String path = currDir.getAbsolutePath();
		String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		this.workbook.write(outputStream);
		this.workbook.close();
	}

}
