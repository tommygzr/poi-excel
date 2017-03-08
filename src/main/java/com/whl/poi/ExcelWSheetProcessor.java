package com.whl.poi;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;

public abstract class ExcelWSheetProcessor<T> {

	public abstract void beforeProcess();

	public abstract List<T> getDataList();

	public abstract void afterProcess();
	
	public abstract void dealException(Exception e);
	
	private Integer sheetIndex = 0;
	private String sheetName;
	private int rowStartIndex = 0;
	private int pageSize = -1;
	private Sheet sheet;

	public Integer getSheetIndex() {
		return sheetIndex;
	}

	public void setSheetIndex(Integer sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public int getRowStartIndex() {
		return rowStartIndex;
	}

	public void setRowStartIndex(int rowStartIndex) {
		this.rowStartIndex = rowStartIndex;
	}

	public int getPageSize() {
		return pageSize;
	}

	public void setPageSize(int pageSize) {
		this.pageSize = pageSize;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public void setSheet(Sheet sheet) {
		this.sheet = sheet;
	}

}
