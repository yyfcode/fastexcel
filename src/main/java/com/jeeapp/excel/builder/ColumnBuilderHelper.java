package com.jeeapp.excel.builder;

/**
 * @author Justice
 * @since 0.0.2
 */
public class ColumnBuilderHelper<B extends ColumnBuilderHelper<B, P>, P extends SheetBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	private final int column;

	protected ColumnBuilderHelper(P parent, int column) {
		super(parent, (short) column);
		this.parent = parent;
		this.column = column;
	}

	/**
	 * 设置列宽
	 */
	public B setColumnWidth(int width) {
		parent.sheet.setColumnWidth(column, width * 256);
		return self();
	}

	/**
	 * 设置换行
	 */
	public B setColumnBreak() {
		parent.sheet.setColumnBreak(column);
		return self();
	}

	/**
	 * 设置折叠
	 */
	public B setColumnGroupCollapsed(boolean collapse) {
		parent.sheet.setColumnGroupCollapsed(column, collapse);
		return self();
	}

	/**
	 * 设置隐藏
	 */
	public B setColumnHidden(boolean hidden) {
		parent.sheet.setColumnHidden(column, hidden);
		return self();
	}
}
