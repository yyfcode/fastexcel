package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author justice
 */
public class CellRangeBuilder<P extends SheetBuilderHelper<P>> extends DataValidationBuilderHelper<CellRangeBuilder<P>, P> {

	private final P parent;

	private final int firstRow;

	private final int lastRow;

	private final int firstCol;

	private final int lastCol;

	protected CellRangeBuilder(P parent, int firstRow, int lastRow, int firstCol, int lastCol) {
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
	public P addMergedRegion() {
		P parent = end();
		parent.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		Cell cell = SheetUtil.getCell(parent.sheet, firstRow, firstCol);
		if (cell != null) {
			parent.setCellStyle(cell);
		}
		return end();
	}

	/**
	 * 添加样式
	 */
	public P addCellStyle() {
		P parent = end();
		for (CellAddress cellAddress : new CellRangeAddress(firstRow, lastRow, firstCol, lastCol)) {
			Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
			if (cell != null) {
				parent.setCellStyle(cell);
			}
		}
		return parent;
	}

	/**
	 * 创建图片
	 */
	public CellRangeBuilder<P> createPicture(byte[] pictureData, int format) {
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(firstRow);
		clientAnchor.setCol1(firstCol);
		clientAnchor.setRow2(lastRow + 1);
		clientAnchor.setCol2(lastCol + 1);
		int pictureIndex = parent.workbook.addPicture(pictureData, format);
		parent.sheet.getDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
		return this;
	}

	/**
	 * 填充未定义的单元格
	 */
	public P fillUndefinedCells() {
		P parent = end();
		for (CellAddress cellAddress : new CellRangeAddress(firstRow, lastRow, firstCol, lastCol)) {
			Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
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
	public CellRangeBuilder<P> setCellComment(String comment, String author, int row2, int col2) {
		parent.createCellComment(comment, author, firstRow, firstCol, row2, col2);
		return this;
	}

	/**
	 * @deprecated use {@link SheetBuilderHelper#matchingRegion(int, int, int, int)} instead.
	 */
	@Deprecated
	public CellRangeBuilder<P> addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		return parent.matchingRegion(firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * @deprecated use {@link CellRangeBuilder#addMergedRegion()} instead.
	 */
	@Deprecated
	public P merge() {
		addMergedRegion();
		return parent;
	}
}
