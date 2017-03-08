package com.whl.poi;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;

public abstract class ExcelRSheetProcessor<T> {

	public abstract void beforeProcess();

	public abstract void process(List<T> list);
	
	public abstract void exceptionHandler(Exception e);

	public abstract void afterProcess();

	private int sheetIndex =0 ;
	private int rowStartIndex = 0;
	private int pageSize = 0;
	private Sheet sheet;

	public int getSheetIndex() {
		return sheetIndex;
	}

	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
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
