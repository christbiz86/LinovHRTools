package com.linov.xlstools.tools;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.linov.xlstools.pojo.SheetPOJO;
import com.linov.xlstools.pojo.XlsReportPOJO;

@Service
public class XlsWriter {
	
	private Workbook workbook;
	private Sheet sheet;
	private Integer rowIndex;
	private Integer maxColumnIndex;
	
	public void writeXls(XlsReportPOJO xlsReportPOJO) throws IOException {
		workbook = new XSSFWorkbook();

		for (SheetPOJO sheetPOJO : xlsReportPOJO.getSheets()) {
			sheet = workbook.createSheet(sheetPOJO.getSheetName());
			
			rowIndex = 0;
			maxColumnIndex = 0;
			
			createRptHeader(sheetPOJO.getRptHeader());
			createGridHeader(sheetPOJO.getGridHeader());
			createContent(sheetPOJO.getContent());
			createFooter();
			
			createFile(xlsReportPOJO.getFileName());
		}
	}

	private void createRptHeader(List<String> rptHeader) {
		CellStyle headerStyle = stylizeRptHeader();
		
		for (int i = 0; i < rptHeader.size(); i++) {
			Row row = sheet.createRow(rowIndex);

			Cell headerCell = row.createCell(0);
			headerCell.setCellValue(rptHeader.get(i));
			headerCell.setCellStyle(headerStyle);
			rowIndex++;
		}
		rowIndex++;
	}

	private CellStyle stylizeRptHeader() {
		CellStyle headerStyle = workbook.createCellStyle();
		
		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 12);
		font.setBold(true);
		headerStyle.setFont(font);
		return headerStyle;
	}

	private void createGridHeader(List<String> gridHeader) {
		Row row = sheet.createRow(rowIndex);

		if (maxColumnIndex < row.getLastCellNum()) {
			maxColumnIndex = row.getLastCellNum() - 1;
		};
		
		for (Integer i = 0; i < gridHeader.size(); i++) {
			CellStyle style = stylizeGridHeader(i, gridHeader.size());
			
			Cell cell = row.createCell(i);
			cell.setCellValue(gridHeader.get(i));
			cell.setCellStyle(style);
		}
		rowIndex++;
	}

	private CellStyle stylizeGridHeader(Integer currentColumn, Integer maxColumn) {
		CellStyle style = workbook.createCellStyle();

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
		
		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 12);
		font.setBold(true);
		style.setFont(font);
		return style;
	}

	private boolean isRightMost(Integer currentColumn, Integer maxColumn) {
		return currentColumn == maxColumn - 1;
	}

	private boolean isLeftMost(Integer currentColumn) {
		return currentColumn == 0;
	}

	private void createContent(List<String[]> content) {
		
		for (Integer i = 0; i < content.size(); i++) {
			Row row = sheet.createRow(rowIndex);
			Integer j = 0;
			
			for (String value : content.get(i)) {
				CellStyle style = stylizeContent(i, j, content.size(), content.get(i).length);
				Cell cell = row.createCell(j);
				cell.setCellValue(value);
				cell.setCellStyle(style);
				j++;
			}
			
			rowIndex++;
		}
	}

	private CellStyle stylizeContent(Integer currentRow, Integer currentColumn, Integer maxRow, Integer maxColumn) {
		CellStyle style = workbook.createCellStyle();

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

		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 12);
		style.setFont(font);
		
		style.setWrapText(true);
		return style;
	}

	private boolean isBottomMost(Integer currentRow, Integer maxRow) {
		return currentRow == maxRow - 1;
	}

	private boolean isTopMost(Integer currentRow) {
		return currentRow == 0;
	}

	private void createFooter() {
		LocalDate dateNow = LocalDate.now();
		rowIndex++;
		Row row = sheet.createRow(rowIndex);
		Cell headerCell = row.createCell(maxColumnIndex);
		headerCell.setCellValue("Jakarta, " + dateNow);
		
		rowIndex += 2;
		row = sheet.createRow(rowIndex);
		headerCell = row.createCell(maxColumnIndex);
		headerCell.setCellValue("(...........................)");
	}
	
	private void createFile(String fileName) throws FileNotFoundException, IOException {
		File currDir = new File(".");
		String path = currDir.getAbsolutePath();
		String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		workbook.write(outputStream);
		workbook.close();
	}

}
