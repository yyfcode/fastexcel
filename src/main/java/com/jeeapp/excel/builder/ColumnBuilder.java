package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author Justice
 * @since 0.0.2
 */
public class ColumnBuilder<B extends ColumnBuilder<B, P>, P extends SheetBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	private final int column;

	protected ColumnBuilder(P parent, int column) {
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

	/**
	 * 添加样式
	 */
	@Override
	public P addCellStyle() {
		P parent = super.addCellStyle();
		int lastRowNum = parent.sheet.getLastRowNum();
		if (lastRowNum > -1) {
			for (CellAddress cellAddress : new CellRangeAddress(0, lastRowNum, column, column)) {
				Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
				if (cell != null) {
					parent.setCellStyle(cell);
				}
			}
		}
		return parent;
	}
}
