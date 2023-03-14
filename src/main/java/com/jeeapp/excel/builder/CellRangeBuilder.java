package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author justice
 */
public class CellRangeBuilder extends CreationBuilder<CellRangeBuilder> {

	private final SheetBuilder parent;

	private final int firstRow;

	private final int lastRow;

	private final int firstCol;

	private final int lastCol;

	protected CellRangeBuilder(SheetBuilder parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
		this.parent = parent;
	}

	/**
	 * 添加合并区域
	 */
	public SheetBuilder addMergedRegion() {
		SheetBuilder parent = super.addCellStyle();
		parent.addMergedRegion(firstRow, lastRow, firstCol, lastCol);
		Cell cell = SheetUtil.getCellWithMerges(parent.sheet, firstRow, firstCol);
		if (cell != null) {
			parent.setCellStyle(cell);
		}
		return parent;
	}

	/**
	 * 设置样式
	 */
	public SheetBuilder setCellStyle() {
		SheetBuilder parent = super.addCellStyle();
		for (CellAddress cellAddress : new CellRangeAddress(firstRow, lastRow, firstCol, lastCol)) {
			Cell cell = SheetUtil.getCellWithMerges(sheet, cellAddress.getRow(), cellAddress.getColumn());
			if (cell != null) {
				parent.setCellStyle(cell);
			}
		}
		return parent;
	}

	/**
	 * 填充未定义的单元格
	 */
	public SheetBuilder fillUndefinedCells() {
		SheetBuilder parent = super.addCellStyle();
		for (CellAddress cellAddress : new CellRangeAddress(firstRow, lastRow, firstCol, lastCol)) {
			Cell cell = SheetUtil.getCellWithMerges(sheet, cellAddress.getRow(), cellAddress.getColumn());
			if (cell == null) {
				parent.createCell(cellAddress);
			}
		}
		return parent;
	}

	/**
	 * @deprecated use {@link CellBuilder#createCellComment(String, String, int, int)} instead.
	 */
	@Deprecated
	public CellRangeBuilder setCellComment(String comment, String author, int row2, int col2) {
		parent.createCellComment(comment, author, firstRow, firstCol, row2, col2);
		return this;
	}

	/**
	 * @deprecated use {@link SheetBuilderHelper#matchingRegion(int, int, int, int)} instead.
	 */
	@Deprecated
	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		return parent.matchingRegion(firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * @deprecated use {@link CellRangeBuilder#addMergedRegion()} instead.
	 */
	@Deprecated
	public SheetBuilder merge() {
		addMergedRegion();
		return parent;
	}
}
