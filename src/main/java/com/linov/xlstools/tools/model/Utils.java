package com.linov.xlstools.tools.model;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Utils {

	private static Integer rowIndex;
	private static List<Integer> maxCharacters;
	
	protected static void generateWorkbook(HSSFWorkbook workbook, List<ExcelSheet> sheets) {
		for (ExcelSheet excelSheet : sheets) {
			Utils.rowIndex = 0;
			Utils.maxCharacters = new ArrayList<Integer>();
			Sheet sheet = workbook.createSheet(excelSheet.getName());
			generateReportHeader(sheet, excelSheet.getHeader());
			generateTable(sheet, excelSheet.getTable());
			generateFooter(sheet, excelSheet.getFooter());
		}
	}
	
	private static void generateReportHeader(Sheet sheet, ExcelHeader header) {
		for (int i = 0; i < header.getTexts().size(); i++) {
			Row row = sheet.createRow(Utils.rowIndex);
			ExcelText text = header.getText(i);
			Cell cell = row.createCell(0);
			setCellValue(text, cell);
			cell.setCellStyle(header.getText(i).getCellStyle());
			Utils.rowIndex++;
		}
	}

	private static void setCellValue(ExcelText text, Cell cell) {
		switch (text.getValueType()) {
		case NUMERIC:
			cell.setCellValue(((Number) text.getValue()).doubleValue());
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

	private static void generateTable(Sheet sheet, ExcelTable table) {
		generateField(sheet, table.getFields());
		generateRecords(sheet, table.getRecords());
		fitTableColumns(sheet, maxCharacters);
	}

	private static void generateField(Sheet sheet, List<ExcelField> fields) {
		Row row = sheet.createRow(Utils.rowIndex);
		Integer cellCharacter = 0;
		
		for (Integer i = 0; i < fields.size(); i++) {
			CellStyle style = fields.get(i).getCellStyle();
			setFieldStyle(style, i, fields.size());
			
			Cell cell = row.createCell(i);
			cell.setCellValue(fields.get(i).getName());
			cell.setCellStyle(style);
			
			DataFormatter df = new DataFormatter();
			String value = df.formatCellValue(cell);
			System.out.println(value);
			cellCharacter = value.length();
			System.out.println(cellCharacter);
			maxCharacters.add(cellCharacter);
		}
		Utils.rowIndex++;
	}

	private static void setFieldStyle(CellStyle style, Integer currentColumn, int maxColumn) {
		style.setAlignment(HorizontalAlignment.CENTER);
		setFieldBorder(style, currentColumn, maxColumn);
	}

	private static void setFieldBorder(CellStyle style, Integer currentColumn, Integer maxColumn) {
		if (isOnlyOneColumn(maxColumn)) {
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
		}
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

	private static boolean isOnlyOneColumn(Integer maxColumn) {
		return maxColumn == 1;
	}

	private static boolean isRightMost(Integer currentColumn, Integer maxColumn) {
		return currentColumn == maxColumn - 1;
	}

	private static boolean isLeftMost(Integer currentColumn) {
		return currentColumn == 0;
	}

	private static void generateRecords(Sheet sheet, List<ExcelRecord> records) {
		Integer i = 0;
		for (ExcelRecord record : records) {
			Row row = sheet.createRow(Utils.rowIndex);

			Integer j = 0;
			List<ExcelText> mapRecord= new ArrayList<ExcelText>(record.get().values());
			
			for (ExcelText text : mapRecord) {
				CellStyle style = text.getCellStyle();
				setRecordsBorder(style, i, j, records.size(), record.size());
				Cell cell = row.createCell(j);
				setCellValue(text, cell);
				cell.setCellStyle(style);
				
				DataFormatter df = new DataFormatter();
				String value = df.formatCellValue(cell);
				System.out.println(value);
				Integer maxCharacter = maxCharacters.get(j);
				Integer cellCharacter = value.length();
				System.out.println(cellCharacter);
				if (maxCharacter < cellCharacter) {
					maxCharacters.set(j, cellCharacter);
				}
				j++;
			}	
			i++;
			Utils.rowIndex++;
		}
	}

	private static CellStyle setRecordsBorder(CellStyle style, Integer currentRow, Integer currentColumn, Integer maxRow, Integer maxColumn) {
		if (isOnlyOneRow(maxRow)) {
			if (isOnlyOneColumn(maxColumn)) {
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			} else {
				if (isLeftMost(currentColumn)) {
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderLeft(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
				} else if (isRightMost(currentColumn, maxColumn)) {
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderRight(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
				} else {
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
				}
			}
		} else if (isOnlyOneColumn(maxColumn)) {
			if (isTopMost(currentRow)) {
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			} else if (isBottomMost(currentRow, maxRow)) {
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
				style.setBorderBottom(BorderStyle.THIN);
			} else {
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			}
		} else if (isTopMost(currentRow)) {
			if (isLeftMost(currentColumn)) {
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
			} else if (isRightMost(currentColumn, maxColumn)) {
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			} else {
				style.setBorderTop(BorderStyle.THIN);
			}
		} else if (isBottomMost(currentRow, maxRow)) { 
			if (isLeftMost(currentColumn)) {
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
			} else if (isRightMost(currentColumn, maxColumn)) {
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
			} else {
				style.setBorderBottom(BorderStyle.THIN);
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

	private static boolean isOnlyOneRow(Integer maxRow) {
		return maxRow == 1;
	}

	private static boolean isBottomMost(Integer currentRow, Integer maxRow) {
		return currentRow == maxRow - 1;
	}

	private static boolean isTopMost(Integer currentRow) {
		return currentRow == 0;
	}

	private static void fitTableColumns(Sheet sheet, List<Integer> maxCharacters) {
		Integer i = 0;
		for (Integer maxCharacter : maxCharacters) {
			System.out.println(maxCharacter);
			int width = ((int)(maxCharacter * 1.14388)) * 256;
			sheet.setColumnWidth(i, width);	
			i++;
		}
	}

	private static void generateFooter(Sheet sheet, ExcelFooter footer) {
		for (int i = 0; i < footer.getTexts().size(); i++) {
			Row row = sheet.createRow(Utils.rowIndex);
			ExcelText text = footer.getText(i);
			Cell cell = row.createCell(0);
			setCellValue(text, cell);
			cell.setCellStyle(footer.getText(i).getCellStyle());
			Utils.rowIndex++;
		}
		Utils.rowIndex++;
	}
}
