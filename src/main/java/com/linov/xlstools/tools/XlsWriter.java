package com.linov.xlstools.tools;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

import com.linov.xlstools.pojo.SheetPOJO;
import com.linov.xlstools.pojo.XlsReportPOJO;
import com.linov.xlstools.tools.model.ExcelField;
import com.linov.xlstools.tools.model.ExcelFooter;
import com.linov.xlstools.tools.model.ExcelHeader;
import com.linov.xlstools.tools.model.ExcelRecord;
import com.linov.xlstools.tools.model.ExcelReport;
import com.linov.xlstools.tools.model.ExcelTable;
import com.linov.xlstools.tools.model.ExcelText;

@Service
public class XlsWriter {
	
	private HSSFWorkbook workbook;
	private Sheet sheet;
	private Integer rowIndex;
	private Integer maxColumnIndex;
	
	public void writeXls(ExcelReport report, XlsReportPOJO xlsReportPOJO) throws IOException {

		for (SheetPOJO sheetPOJO : xlsReportPOJO.getSheets()) {
			this.sheet = this.workbook.createSheet(sheetPOJO.getSheetName());
			
			this.rowIndex = 0;
			this.maxColumnIndex = 0;

			setExcelReport(report, sheetPOJO);
			
			createFile(report.getFileName());
		}
	}

	private void setExcelReport(ExcelReport report, SheetPOJO sheetPOJO) {
		List<String> rptHeader = new ArrayList<String>();
		List<String> gridHeader = new ArrayList<String>();
		List<String[]> content = new ArrayList<String[]>();
		List<String> footer = new ArrayList<String>();
		////////////////
		ExcelHeader header = report.getHeader();
		for (String string : sheetPOJO.getRptHeader()) {
			header.addText(string);
			System.out.println("header size: " + header.getTexts().size());
		}
		for (ExcelText text : header.getTexts()) {
			rptHeader.add(text.getValue().toString());
		}
//		//////////////////
		ExcelTable table = report.getTable();
		for (String string : sheetPOJO.getGridHeader()) {
			table.addField(string);
			System.out.println("field size: " + table.getFields().size());
		}
		for (ExcelField field : table.getFields()) {
			gridHeader.add(field.getName());
		}
		///////////////
		for (String[] strings : sheetPOJO.getContent()) {
			List<Object> convStrings = Arrays.asList(strings);
			System.out.println(table.getFields().size() + " : " + table.getRecords().size());
			table.addRecord(convStrings);
		}
		for (ExcelRecord record : table.getRecords()) {
			System.out.println(record.get().size());
			System.out.println(((ExcelText) ((record.get().values().toArray())[0])).getValue());
			ExcelText[] temp1 = (record.get().values().toArray(new ExcelText[record.get().size()]));
			String[] temp2 = new String[temp1.length];
			int j = 0;
			for (ExcelText text : temp1) {
				temp2[j] = (String) text.getValue();
			}
//			System.out.println(Arrays.toString(record.get().values().toArray(new String[record.get().size()])));
			
			content.add(temp2);
		}
		/////////////
		ExcelFooter foot = report.getFooter();
		for (String string : sheetPOJO.getFooter()) {
			foot.addText(string);
		}
		for (ExcelText text : foot.getTexts()) {
			footer.add(text.getValue().toString());
		}
		///////////
		
		createRptHeader(rptHeader);
		createGridHeader(gridHeader);
		createContent(content);
		createFooter(footer);
	}

	private void createRptHeader(List<String> rptHeader) {
		CellStyle headerStyle = stylizeRptHeader();
		
		for (int i = 0; i < rptHeader.size(); i++) {
			Row row = this.sheet.createRow(this.rowIndex);

			Cell headerCell = row.createCell(0);
			headerCell.setCellValue(rptHeader.get(i));
			headerCell.setCellStyle(headerStyle);
			this.rowIndex++;
		}
	}

	private CellStyle stylizeRptHeader() {
		CellStyle headerStyle = this.workbook.createCellStyle();
		
		HSSFFont font = ((HSSFWorkbook) this.workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 12);
		font.setBold(true);
		headerStyle.setFont(font);
		return headerStyle;
	}

	private void createGridHeader(List<String> gridHeader) {
		Row row = this.sheet.createRow(this.rowIndex);

		if (this.maxColumnIndex < row.getLastCellNum()) {
			this.maxColumnIndex = row.getLastCellNum() - 1;
		};
		
		for (Integer i = 0; i < gridHeader.size(); i++) {
			CellStyle style = stylizeGridHeader(i, gridHeader.size());
			
			Cell cell = row.createCell(i);
			cell.setCellValue(gridHeader.get(i));
			cell.setCellStyle(style);
		}
		this.rowIndex++;
	}

	private CellStyle stylizeGridHeader(Integer currentColumn, Integer maxColumn) {
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
		
		HSSFFont font = ((HSSFWorkbook) this.workbook).createFont();
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

	private CellStyle stylizeContent(Integer currentRow, Integer currentColumn, Integer maxRow, Integer maxColumn) {
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

	private boolean isBottomMost(Integer currentRow, Integer maxRow) {
		return currentRow == maxRow - 1;
	}

	private boolean isTopMost(Integer currentRow) {
		return currentRow == 0;
	}

	private void createFooter(List<String> footer) {
		CellStyle footerStyle = stylizeRptHeader();
		
		for (int i = 0; i < footer.size(); i++) {
			Row row = this.sheet.createRow(this.rowIndex);

			Cell footerCell = row.createCell(0);
			footerCell.setCellValue(footer.get(i));
			footerCell.setCellStyle(footerStyle);
			this.rowIndex++;
		}
		this.rowIndex++;
		
//		LocalDate dateNow = LocalDate.now();
//		this.rowIndex++;
//		Row row = this.sheet.createRow(this.rowIndex);
//		Cell headerCell = row.createCell(this.maxColumnIndex);
//		headerCell.setCellValue("Jakarta, " + dateNow);
//		
//		this.rowIndex += 2;
//		row = this.sheet.createRow(this.rowIndex);
//		headerCell = row.createCell(this.maxColumnIndex);
//		headerCell.setCellValue("(...........................)");
	}
	
	private void createFile(String fileName) throws FileNotFoundException, IOException {
		File currDir = new File(".");
		String path = currDir.getAbsolutePath();
		String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";

		FileOutputStream outputStream = new FileOutputStream(fileLocation);
		this.workbook.write(outputStream);
		this.workbook.close();
	}

}
