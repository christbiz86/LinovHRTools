package com.linov.xlstools.pojo;

public class RangePOJO {

	private Integer startRow;
	private Integer endRow;
	private Integer startColumn;
	private Integer endColumn;
	
	public RangePOJO(String startCell, String endCell) {
		startRow = getStartRow(startCell);
		endRow = getEndRow(endCell);
		startColumn = getStartColumn(startCell);
		endColumn = getEndColumn(endCell);
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
	
	private Integer getStartRow(String startCell) {
		Integer startIndex = getStartIndex(startCell);
		Integer row = Integer.parseInt(startCell.substring(startIndex));
		return row - 1;
	}

	private Integer getStartColumn(String startCell) {
		Integer startIndex = getStartIndex(startCell);
		String Column = startCell.substring(0, startIndex);
		return excelColumnNameToNumber(Column);
	}
	
	private Integer getEndRow(String endCell) {
		Integer startIndex = getStartIndex(endCell);
		Integer row = Integer.parseInt(endCell.substring(startIndex));
		return row - 1;
	}
	
	private Integer getEndColumn(String endCell) {
		Integer startIndex = getStartIndex(endCell);
		String Column = endCell.substring(0, startIndex);
		return excelColumnNameToNumber(Column);
	}
	
	private Integer excelColumnNameToNumber(String ColumnName) {
	    ColumnName = ColumnName.toUpperCase();
	    Integer sum = 0;
	    for (Integer i = 0; i < ColumnName.length(); i++)
	    {
	        sum *= 26;
	        sum += (ColumnName.charAt(i) - 'A');
	    }
	    return sum;
	}

	public Integer getStartRow() {
		return startRow;
	}

	public Integer getEndRow() {
		return endRow;
	}

	public Integer getStartColumn() {
		return startColumn;
	}

	public Integer getEndColumn() {
		return endColumn;
	}
	
	
}
