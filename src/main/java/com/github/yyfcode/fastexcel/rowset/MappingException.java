package com.github.yyfcode.fastexcel.rowset;

import com.github.yyfcode.fastexcel.model.Row;

/**
 * @author Justice
 */
public class MappingException extends Exception implements MappingErrors {

	private static final long serialVersionUID = 2094426361433877913L;

	private final Row row;

	public MappingException(Row row) {
		this.row = row;
	}

	@Override
	public Row getRow() {
		return row;
	}

	@Override
	public void addError(int colNum, String errMsg) {
		row.addComment(colNum, errMsg);
	}

	@Override
	public boolean hasErrors() {
		return row.hasComments();
	}
}
