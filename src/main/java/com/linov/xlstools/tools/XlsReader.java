package com.linov.xlstools.tools;

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;

import com.linov.xlstools.pojo.RangePOJO;

@Service
public class XlsReader {

	public List<Map<String, Object>> readXls(InputStream file) throws IOException {
		return readXls(file, "");
	}

	public List<Map<String, Object>> readXls(InputStream file, String startCell, String endCell) throws IOException {
		return readXls(file, "", startCell, endCell);
	}

	public List<Map<String, Object>> readXls(InputStream file, String sheetName) throws IOException {
		Workbook workbook = createWorkbook(file);
		
		Sheet sheet = getSheet(sheetName, workbook);
		if (isNull(sheet)) {
			return new ArrayList<>();
		}
		
		List<Map<String, Object>> records = new ArrayList<Map<String,Object>>();
		
		readCells(sheet, records);

		workbook.close();
		return records;
	}

	private Workbook createWorkbook(InputStream file) {
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(file);
		}
		catch (IOException e) {
			throw new IllegalArgumentException("File does not have a standard excel extension(.xls or .xlsx");
		}
		return workbook;
	}
	
	public List<Map<String, Object>> readXls(InputStream file, String sheetName, String startCell, String endCell) throws IOException {
		Workbook workbook = WorkbookFactory.create(file);
		
		Sheet sheet = getSheet(sheetName, workbook);
		if (isNull(sheet)) {
			return new ArrayList<>();
		}
		
		List<Map<String, Object>> records = new ArrayList<Map<String,Object>>();
		
		RangePOJO range = new RangePOJO(startCell, endCell);
		
		readCells(sheet, records, range);

		workbook.close();
		return records;
	}

	private boolean isNull(Object object) {
		return object == null;
	}

	private void readCells(Sheet sheet, List<Map<String, Object>> records) {
		List<String> keys = new ArrayList<String>();
		for (Row row : sheet) {
			if (isNull(row)) {
				continue;
			}
			Map<String, Object> record= new HashMap<String, Object>();
			List<Object>values = new ArrayList<Object>();
			
		    parseCells(keys, row, values);
		    addRecord(records, keys, record, values);
		    
			if (!keys.isEmpty() && hasReachEndOfRecord(row)) {
				break;
			}
		}
	}

	private void readCells(Sheet sheet, List<Map<String, Object>> records, RangePOJO range) {
		List<String> keys = new ArrayList<String>();
		for (int i = range.getStartRow(); i <= range.getEndRow(); i++) {
			Row row = sheet.getRow(i);
			if (isNull(row)) {
				continue;
			}
			Map<String, Object> record= new HashMap<String, Object>();
			List<Object>values = new ArrayList<Object>();
			
		    parseCells(keys, row, values, range);
		    addRecord(records, keys, record, values);
		    
			if (!keys.isEmpty() && hasReachEndOfRecord(row)) {
				break;
			}
		}
	}

	private Sheet getSheet(String sheetName, Workbook workbook) {
		Sheet sheet;
		if (isNull(sheetName) || sheetName.isEmpty()) {
			sheet = workbook.getSheetAt(0);
		} else if (workbook.getSheetIndex(sheetName) > 0){
			sheet = workbook.getSheet(sheetName);
		} else {
			return null;
		}
		return sheet;
	}

	private void parseCells(List<String> keys, Row row, List<Object> values) {
		for (Cell cell : row) {
			if (isNull(cell)) {
				continue;
			}
			if (isGridHeader(cell)) {
				keys.add(cell.getStringCellValue());
			}
			else if (!keys.isEmpty()){
				values.add(getValue(cell));
			}
		}
	}

	private void parseCells(List<String> keys, Row row, List<Object> values, RangePOJO range) {
		for (int i = range.getStartColumn(); i <= range.getEndColumn(); i++) {
			Cell cell = row.getCell(i);
			if (isNull(cell)) {
				continue;
			}
			if (isGridHeader(cell)) {
				keys.add(cell.getStringCellValue());
			}
			else if (!keys.isEmpty()){
				values.add(getValue(cell));
			}
		}
	}

	private void addRecord(List<Map<String, Object>> records, List<String> keys, Map<String, Object> record,
			List<Object> values) {
		if (!values.isEmpty()) {
			for (Integer j = 0; j < keys.size(); j++) {
				record.put(keys.get(j), values.get(j));
			}
			records.add(record);
		}
	}

	private boolean hasReachEndOfRecord(Row row) {
		return row.getCell(0).getCellStyle().getBorderBottom() != BorderStyle.NONE && row.getCell(0).getCellStyle().getBorderTop() == BorderStyle.NONE;
	}

	private Object getValue(Cell cell) {
		switch (cell.getCellType()) {
		    case STRING: 
		    	return cell.getRichStringCellValue().getString();
		    case NUMERIC: 
		    	if (DateUtil.isCellDateFormatted(cell)) {
		    		LocalDateTime ldt = LocalDateTime.ofInstant(cell.getDateCellValue().toInstant(), ZoneId.systemDefault());
		    	    return ldt;
		    	} else {
		    	    return cell.getNumericCellValue();
		    	}
		    case BOOLEAN: 
		    	return cell.getBooleanCellValue();
		    case FORMULA: 
		    	return getFormulaResult(cell);
		    default: 
		    	return "";
		}
	}

	private Object getFormulaResult(Cell cell) {
		switch(cell.getCachedFormulaResultType()) {
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					LocalDateTime ldt = LocalDateTime.ofInstant(cell.getDateCellValue().toInstant(), ZoneId.systemDefault());
				    return ldt;
				} else {
				    return cell.getNumericCellValue();
				}
			case STRING:
				return cell.getRichStringCellValue().getString();
			case BOOLEAN:
				return cell.getBooleanCellValue();
			default:
				return "";
		}
	}

	private boolean isGridHeader(Cell cell) {
		return cell.getCellStyle().getBorderBottom() != BorderStyle.NONE && cell.getCellStyle().getBorderTop() != BorderStyle.NONE;
	}

}