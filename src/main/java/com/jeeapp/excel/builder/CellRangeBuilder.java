package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author justice
 */
public class CellRangeBuilder<P extends RowBuilderHelper<P>> extends DataValidationBuilder<CellRangeBuilder<P>, P> {

	private final P parent;

	private final int firstRow;

	private final int lastRow;

	private final int firstCol;

	private final int lastCol;

	private Object value;

	protected CellRangeBuilder(P parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
		this.parent = parent;
	}

	public CellBuilder<P> addMergedRegion() {
		parent.createCell(lastRow, lastCol, null);
		parent.sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		return parent.matchingCell(new CellAddress(firstRow, firstCol));
	}

	public CellRangeBuilder<P> addPicture(byte[] pictureData, int format) {
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(firstRow);
		clientAnchor.setCol1(firstCol);
		clientAnchor.setRow2(lastRow + 1);
		clientAnchor.setCol2(lastCol + 1);
		int pictureIndex = parent.workbook.addPicture(pictureData, format);
		parent.sheet.getDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
		return this;
	}

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
	 * @deprecated use {@link CellBuilder#setCommentText(String)} instead.
	 */
	@Deprecated
	public CellRangeBuilder<P> setCellComment(String comment, String author, int row2, int col2) {
		parent.createCellComment(comment, author, firstRow, firstCol, row2, col2);
		return this;
	}

	/**
	 * @deprecated use {@link RowBuilderHelper#matchingRegion(int, int, int, int)} instead.
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
		return addMergedRegion().end();
	}
}
