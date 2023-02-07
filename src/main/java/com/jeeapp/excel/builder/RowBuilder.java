package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author Justice
 */
public class RowBuilder<B extends RowBuilder<B, P>, P extends SheetBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	private final int rowNum;

	protected RowBuilder(P parent, int rowNum) {
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
	 * 添加样式
	 */
	@Override
	public P addCellStyle() {
		P parent = super.addCellStyle();
		Row row = parent.sheet.getRow(rowNum);
		if (row != null) {
			short lastCellNum = row.getLastCellNum();
			if (lastCellNum > -1) {
				for (CellAddress cellAddress : new CellRangeAddress(rowNum, rowNum, 0, lastCellNum)) {
					Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
					if (cell != null) {
						parent.setCellStyle(cell);
					}
				}
			}
		}
		return parent;
	}
}
