package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author Justice
 * @since 0.0.2
 */
public class ColumnBuilder extends CellStyleBuilder<ColumnBuilder, SheetBuilder> {

	private final SheetBuilder parent;

	private final int column;

	protected ColumnBuilder(SheetBuilder parent, int column) {
		super(parent, (short) column);
		this.parent = parent;
		this.column = column;
	}

	/**
	 * 设置列宽
	 */
	public ColumnBuilder setColumnWidth(int width) {
		parent.sheet.setColumnWidth(column, width * 256);
		return self();
	}

	/**
	 * 设置换行
	 */
	public ColumnBuilder setColumnBreak() {
		parent.sheet.setColumnBreak(column);
		return self();
	}

	/**
	 * 设置折叠
	 */
	public ColumnBuilder setColumnGroupCollapsed(boolean collapse) {
		parent.sheet.setColumnGroupCollapsed(column, collapse);
		return self();
	}

	/**
	 * 设置隐藏
	 */
	public ColumnBuilder setColumnHidden(boolean hidden) {
		parent.sheet.setColumnHidden(column, hidden);
		return self();
	}

	/**
	 * 给匹配列设置样式
	 */
	public SheetBuilder setCellStyle() {
		SheetBuilder parent = super.addCellStyle();
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
