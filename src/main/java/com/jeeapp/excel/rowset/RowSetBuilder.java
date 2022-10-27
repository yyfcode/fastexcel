package com.jeeapp.excel.rowset;

import com.jeeapp.excel.model.Row;

/**
 * @author Justice
 */
public final class RowSetBuilder {

	private int sheetIndex;

	private String sheetName;

	private int lastRowNum;

	private Row row;

	private RowSetBuilder() {}

	private RowSetBuilder(RowSet rowSet) {
		this.sheetIndex = rowSet.getSheetIndex();
		this.sheetName = rowSet.getSheetName();
		this.lastRowNum = rowSet.getLastRowNum();
		this.row = rowSet.getRow();
	}

	public static RowSetBuilder builder() {
		return new RowSetBuilder();
	}

	public static RowSetBuilder builder(RowSet rowSet) {
		return new RowSetBuilder(rowSet);
	}

	public RowSetBuilder withSheet(int sheetIndex, String sheetName) {
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
		return this;
	}

	public RowSetBuilder withLastRowNum(int lastRowNum) {
		this.lastRowNum = lastRowNum;
		return this;
	}

	public RowSetBuilder withRow(int rowNum, String[] cellValues) {
		this.row = new Row(rowNum, cellValues);
		return this;
	}

	public RowSetBuilder withNullRow() {
		this.row = null;
		return this;
	}

	public RowSet build() {
		RowSetImpl rowSetImpl = new RowSetImpl();
		rowSetImpl.setSheetIndex(sheetIndex);
		rowSetImpl.setSheetName(sheetName);
		rowSetImpl.setLastRowNum(lastRowNum);
		rowSetImpl.setRow(row);
		return rowSetImpl;
	}
}
