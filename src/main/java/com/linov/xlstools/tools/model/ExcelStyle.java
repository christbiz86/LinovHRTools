package com.linov.xlstools.tools.model;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class ExcelStyle {

	private CellStyle cellStyle;
	private CreationHelper creationHelper;

	protected ExcelStyle(HSSFWorkbook workbook) {
		this.creationHelper = workbook.getCreationHelper();
		this.cellStyle = workbook.createCellStyle();
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	public short getIndex() {
		return this.cellStyle.getIndex();
	}

	public void setDataFormat(String dateFormat) {
		short fmt = this.creationHelper.createDataFormat().getFormat(dateFormat);
		this.cellStyle.setDataFormat(fmt);
	}

	public short getDataFormat() {
		return this.cellStyle.getDataFormat();
	}

	public String getDataFormatString() {
		return this.cellStyle.getDataFormatString();
	}

	public void setHidden(boolean hidden) {
		this.cellStyle.setHidden(hidden);
	}

	public boolean getHidden() {
		return this.cellStyle.getHidden();
	}

	public void setLocked(boolean locked) {
		this.cellStyle.setLocked(locked);
	}

	public boolean getLocked() {
		return this.cellStyle.getLocked();
	}

	public void setQuotePrefixed(boolean quotePrefix) {
		this.cellStyle.setQuotePrefixed(quotePrefix);
	}

	public boolean getQuotePrefixed() {
		return this.cellStyle.getQuotePrefixed();
	}

	public void setAlignment(HorizontalAlignment align) {
		this.cellStyle.setAlignment(align);
	}

	public HorizontalAlignment getAlignment() {
		return this.cellStyle.getAlignment();
	}

	@Deprecated
	public HorizontalAlignment getAlignmentEnum() {
		return this.cellStyle.getAlignmentEnum();
	}

	public void setWrapText(boolean wrapped) {
		this.cellStyle.setWrapText(wrapped);
	}

	public boolean getWrapText() {
		return this.cellStyle.getWrapText();
	}

	public void setVerticalAlignment(VerticalAlignment align) {
		this.cellStyle.setVerticalAlignment(align);
	}

	public VerticalAlignment getVerticalAlignment() {
		return this.cellStyle.getVerticalAlignment();
	}

	@Deprecated
	public VerticalAlignment getVerticalAlignmentEnum() {
		return this.cellStyle.getVerticalAlignmentEnum();
	}

	public void setRotation(short rotation) {
		this.cellStyle.setRotation(rotation);
	}

	public short getRotation() {
		return this.cellStyle.getRotation();
	}

	public void setIndention(short indent) {
		this.cellStyle.setIndention(indent);
	}

	public short getIndention() {
		return this.cellStyle.getIndention();
	}

	public void setFillPattern(FillPatternType fp) {
		this.cellStyle.setFillPattern(fp);
	}

	public FillPatternType getFillPattern() {
		return this.cellStyle.getFillPattern();
	}

	@Deprecated
	public FillPatternType getFillPatternEnum() {
		return this.cellStyle.getFillPatternEnum();
	}

	public void setFillBackgroundColor(short bg) {
		this.cellStyle.setFillBackgroundColor(bg);
	}

	public short getFillBackgroundColor() {
		return this.cellStyle.getFillBackgroundColor();
	}

	public Color getFillBackgroundColorColor() {
		return this.cellStyle.getFillBackgroundColorColor();
	}

	public void setFillForegroundColor(short bg) {
		this.cellStyle.setFillForegroundColor(bg);
	}

	public short getFillForegroundColor() {
		return this.cellStyle.getFillForegroundColor();
	}

	public Color getFillForegroundColorColor() {
		return this.cellStyle.getFillForegroundColorColor();
	}

	public void cloneStyleFrom(CellStyle source) {
		this.cellStyle.cloneStyleFrom(source);
	}

	public void setShrinkToFit(boolean shrinkToFit) {
		this.cellStyle.setShrinkToFit(shrinkToFit);
	}

	public boolean getShrinkToFit() {
		return this.cellStyle.getShrinkToFit();
	}

}
