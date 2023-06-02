package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author Justice
 */
public class RowBuilder extends CellStyleBuilder<RowBuilder, SheetBuilder> {

	private final SheetBuilder parent;

	private final int rowNum;

	protected RowBuilder(SheetBuilder parent, int rowNum) {
		super(parent, rowNum);
		this.parent = parent;
		this.rowNum = rowNum;
	}

	/**
	 * 设置行高
	 */
	public RowBuilder setRowHeight(int height) {
		Row row = parent.sheet.getRow(rowNum);
		if (row != null) {
			row.setHeightInPoints(height);
		}
		return self();
	}

	/**
	 * 设置换行
	 */
	public RowBuilder setRowBreak() {
		parent.sheet.setRowBreak(rowNum);
		return self();
	}

	/**
	 * 给匹配行设置样式
	 */
	public SheetBuilder setCellStyle() {
		SheetBuilder parent = super.addCellStyle();
		Row row = parent.sheet.getRow(rowNum);
		if (row != null && row.getLastCellNum() > -1) {
			for (CellAddress cellAddress : new CellRangeAddress(rowNum, rowNum, 0, row.getLastCellNum())) {
				Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
				if (cell != null) {
					parent.setCellStyle(cell);
				}
			}
		}
		return parent;
	}
}
