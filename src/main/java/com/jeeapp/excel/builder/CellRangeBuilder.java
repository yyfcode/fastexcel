package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author justice
 */
public class CellRangeBuilder extends PictureBuilder<CellRangeBuilder> {

	private final SheetBuilder parent;

	public CellRangeBuilder(SheetBuilder parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.parent = parent;
	}

	/**
	 * @deprecated use {@link CellBuilder#setCellValue(Object)} instead.
	 */
	@Deprecated
	public CellRangeBuilder setCellValue(Object value) {
		parent.createCell(region.getFirstRow(), region.getFirstColumn(), value);
		return this;
	}

	public SheetBuilder fillUndefinedCells() {
		for (CellAddress cellAddress : region) {
			Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
			if (cell == null) {
				parent.createCell(cellAddress);
			}
		}
		return end();
	}

	public CellBuilder addMergedRegion() {
		parent.sheet.addMergedRegion(region);
		return end().matchingCell(new CellAddress(region.getFirstRow(), region.getFirstColumn()));
	}

	public CellRangeBuilder matchingRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
		return end().matchingRegion(firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * @deprecated use {@link CellBuilder#setCommentText(String)} instead.
	 */
	@Deprecated
	public CellRangeBuilder setCellComment(String comment, String author, int row2, int col2) {
		parent.createCellComment(comment, author, region.getFirstRow(), region.getFirstColumn(), row2, col2);
		return this;
	}

	/**
	 * @deprecated use {@link CellRangeBuilder#matchingRegion(int, int, int, int)} instead.
	 */
	@Deprecated
	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		return end().matchingRegion(firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * @deprecated use {@link CellRangeBuilder#addMergedRegion()} instead.
	 */
	@Deprecated
	public SheetBuilder merge() {
		return addMergedRegion().end();
	}
}
