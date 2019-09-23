package com.linov.xlstools.xlstools;

import java.io.IOException;
import java.io.InputStream;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class XlsReader {

	public List<Map<String, Object>> readXls(InputStream file, String startCell, String endCell) throws IOException {
		Workbook workbook = new XSSFWorkbook(file);
		
		Sheet sheet = workbook.getSheetAt(0);
		 
		List<Map<String, Object>> records = new ArrayList<Map<String,Object>>();
		List<String> keys = new ArrayList<String>();
		
		Integer startRow = getStartRow(startCell);
		Integer endRow = getStartRow(endCell);
		Integer startColumn = getStartRow(startCell);
		Integer endColumn = getStartRow(endCell);
		
		for (int i = startRow; i < endRow; i++) {
			Row row = sheet.getRow(i);
			Map<String, Object> record= new HashMap<String, Object>();
			List<Object>values = new ArrayList<Object>();
			
		    parseCells(keys, row, values, startColumn, endColumn);
		    addRecord(records, keys, record, values);
		    
			if (hasReachEndOfRecord(row)) {
				break;
			}
		}

		workbook.close();
		return records;
	}

	private void parseCells(List<String> keys, Row row, List<Object> values, Integer startColumn, Integer endColumn) {

		for (int i = startColumn; i < endColumn; i++) {
			Cell cell = row.getCell(i);
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
		    	    return cell.getDateCellValue();
		    	} else {
		    	    return cell.getNumericCellValue();
		    	}
		    case BOOLEAN: 
		    	return cell.getBooleanCellValue();
		    case FORMULA: 
		    	return cell.getCellFormula();
		    default: 
		    	return " ";
		}
	}

	private boolean isGridHeader(Cell cell) {
		return cell.getCellStyle().getBorderBottom() != BorderStyle.NONE && cell.getCellStyle().getBorderTop() != BorderStyle.NONE;
	}

	private Integer getStartIndex(String startCell) {
		String numbers = "0123456789";
		for(int i = 0; i < numbers.length(); i++){
		    if(startCell.indexOf(numbers.charAt(i)) > -1){
				return startCell.indexOf(numbers.charAt(i));
		    }
		}
		return null;
	}
	
	public Integer getStartRow(String startCell) {
		Integer startIndex = getStartIndex(startCell);
		System.out.println(startIndex);
		Integer row = Integer.parseInt(startCell.substring(startIndex));
		return row;
	}

	public Integer getStartColumn(String startCell) {
		Integer startIndex = getStartIndex(startCell);
		String Column = startCell.substring(0, startIndex);
		return excelColumnNameToNumber(Column);
	}
	
	public Integer getEndRow(String endCell) {
		Integer startIndex = getStartIndex(endCell);
		Integer row = Integer.parseInt(endCell.substring(startIndex));
		return row;
	}
	
	public Integer getEndColumn(String endCell) {
		Integer startIndex = getStartIndex(endCell);
		String Column = endCell.substring(0, startIndex);
		return excelColumnNameToNumber(Column);
	}
	
	public Integer excelColumnNameToNumber(String ColumnName) {
	    ColumnName = ColumnName.toUpperCase();

	    Integer sum = 0;

	    for (Integer i = 0; i < ColumnName.length(); i++)
	    {
	        sum *= 26;
	        sum += (ColumnName.charAt(i) - 'A' + 1);
	    }

	    return sum;
	}
}
