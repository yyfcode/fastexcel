package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author Justice
 * @since 0.0.2
 */
public class RowBuilderHelper<B extends RowBuilderHelper<B, P>, P extends SheetBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	private final int rowNum;

	protected RowBuilderHelper(P parent, int rowNum) {
		super(parent, rowNum);
		this.parent = parent;
		this.rowNum = rowNum;
	}

	/**
	 * 设置行高
	 */
	public B setRowHeight(int height) {
		Row row = parent.sheet.getRow(rowNum);
		if (row != null) {
			row.setHeightInPoints(height);
		}
		return self();
	}

	/**
	 * 设置换行
	 */
	public B setRowBreak() {
		parent.sheet.setRowBreak(rowNum);
		return self();
	}

	/**
	 * 设置折叠
	 */
	public B setRowGroupCollapsed(boolean collapse) {
		parent.sheet.setRowGroupCollapsed(rowNum, collapse);
		return self();
	}
}
