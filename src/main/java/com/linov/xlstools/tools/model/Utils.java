package com.linov.xlstools.tools.model;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Utils {

	private Integer rowIndex;
	private HSSFWorkbook workbook;

	public Utils() {
		this.rowIndex = 0;
	}
	
	public void generateWorkbook(HSSFWorkbook workbook, List<ExcelSheet> sheets) {
		for (ExcelSheet excelSheet : sheets) {
			Sheet sheet = workbook.createSheet(excelSheet.getName());
			generateReportHeader(sheet, excelSheet.getHeader());
			generateTable(sheet, excelSheet.getTable());
			generateFooter(sheet, excelSheet.getFooter());
		}
	}
	
	public void generateReportHeader(Sheet sheet, ExcelHeader header) {
		for (int i = 0; i < header.getTexts().size(); i++) {
			Row row = sheet.createRow(this.rowIndex);
			ExcelText text = header.getText(i);
			Cell cell = row.createCell(0);
			setCellValue(text, cell);
			cell.setCellStyle(header.getText(i).getCellStyle());
			this.rowIndex++;
		}
	}

	private void setCellValue(ExcelText text, Cell cell) {
		switch (text.getValueType()) {
		case NUMERIC:
			cell.setCellValue((Double) text.getValue());
			break;
		case DATE:
			cell.setCellValue((Date) text.getValue());
			break;
		case BOOLEAN:
			cell.setCellValue((Boolean) text.getValue());
			break;
		default:
			cell.setCellValue((String) text.getValue());
			break;
		}
	}

	private void generateTable(Sheet sheet, ExcelTable table) {
		generateField(sheet, table.getFields());
		generateRecords(sheet, table.getRecords());
	}

	public void generateField(Sheet sheet, List<ExcelField> fields) {
		Row row = sheet.createRow(this.rowIndex);

		for (Integer i = 0; i < fields.size(); i++) {
			CellStyle style = fields.get(i).getCellStyle();
			setFieldBorder(style, i, fields.size());
			
			Cell cell = row.createCell(i);
			cell.setCellValue(fields.get(i).getName());
			cell.setCellStyle(style);
		}
		this.rowIndex++;
	}

	public void setFieldBorder(CellStyle style, Integer currentColumn, Integer maxColumn) {
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
	}

	public boolean isRightMost(Integer currentColumn, Integer maxColumn) {
		return currentColumn == maxColumn - 1;
	}

	public boolean isLeftMost(Integer currentColumn) {
		return currentColumn == 0;
	}

	public void generateRecords(Sheet sheet, List<ExcelRecord> records) {
		for (ExcelRecord record : records) {
			Row row = sheet.createRow(this.rowIndex);
			
			for (Integer j = 0; j < record.size(); j++) {
				List<ExcelText> mapRecord= new ArrayList<ExcelText>(record.get().values());
				
				for (ExcelText text : mapRecord) {
					CellStyle style = text.getCellStyle();
					setRecordsBorder(style, row.getRowNum(), j, records.size(), record.size());
					Cell cell = row.createCell(j);
					setCellValue(text, cell);
					cell.setCellStyle(style);
				}
				j++;
			}
			this.rowIndex++;
		}
	}

	public CellStyle setRecordsBorder(CellStyle style, Integer currentRow, Integer currentColumn, Integer maxRow, Integer maxColumn) {
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
		
		style.setWrapText(true);
		return style;
	}

	public boolean isBottomMost(Integer currentRow, Integer maxRow) {
		return currentRow == maxRow - 1;
	}

	public boolean isTopMost(Integer currentRow) {
		return currentRow == 0;
	}

	public void generateFooter(Sheet sheet, ExcelFooter footer) {
		for (int i = 0; i < footer.getTexts().size(); i++) {
			Row row = sheet.createRow(this.rowIndex);
			ExcelText text = footer.getText(i);
			Cell cell = row.createCell(0);
			setCellValue(text, cell);
			cell.setCellStyle(footer.getText(i).getCellStyle());
			this.rowIndex++;
		}
		this.rowIndex++;
	}
}
