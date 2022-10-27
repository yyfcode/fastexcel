package com.jeeapp.excel.rowset;

import com.jeeapp.excel.model.Comment;
import com.jeeapp.excel.model.Row;

/**
 * @author Justice
 */
public class MappingException extends Exception implements MappingErrors {

	private static final long serialVersionUID = 2094426361433877913L;

	private final Row row;

	public MappingException(Row row) {
		this.row = row;
	}

	public MappingException(Row row, Throwable cause) {
		super(cause);
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

	@Override
	public String getMessage() {
		StringBuilder sb = new StringBuilder("Row ")
			.append(row.getRowNum())
			.append(" has ")
			.append(row.getComments().size()).append(" errors :");
		for (Comment comment : row.getComments()) {
			sb.append('\n')
				.append("Column ")
				.append(comment.getColNum())
				.append(" [")
				.append(row.getCellValues()[comment.getColNum()])
				.append("] ")
				.append(comment.getText());
		}
		return sb.toString();
	}
}
