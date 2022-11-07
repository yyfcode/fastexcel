package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

/**
 * @author justice
 */
public class CellRangeBuilder {

	private final SheetBuilder parent;

	private final CellRangeAddress region;

	public CellRangeBuilder(SheetBuilder parent, CellRangeAddress region) {
		this.parent = parent;
		this.region = region;
	}

	public CellRangeBuilder setBorder(BorderStyle border) {
		RegionUtil.setBorderLeft(border, region, parent.sheet);
		RegionUtil.setBorderBottom(border, region, parent.sheet);
		RegionUtil.setBorderRight(border, region, parent.sheet);
		RegionUtil.setBorderTop(border, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBorderColor(int color) {
		RegionUtil.setLeftBorderColor(color, region, parent.sheet);
		RegionUtil.setBottomBorderColor(color, region, parent.sheet);
		RegionUtil.setRightBorderColor(color, region, parent.sheet);
		RegionUtil.setTopBorderColor(color, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBorderLeft(BorderStyle border) {
		RegionUtil.setBorderLeft(border, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBorderBottom(BorderStyle border) {
		RegionUtil.setBorderBottom(border, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBorderRight(BorderStyle border) {
		RegionUtil.setBorderRight(border, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBorderTop(BorderStyle border) {
		RegionUtil.setBorderTop(border, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setLeftBorderColor(int color) {
		RegionUtil.setLeftBorderColor(color, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBottomBorderColor(int color) {
		RegionUtil.setBottomBorderColor(color, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setRightBorderColor(int color) {
		RegionUtil.setRightBorderColor(color, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setTopBorderColor(int color) {
		RegionUtil.setTopBorderColor(color, region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setLeftBorderColor(IndexedColors color) {
		RegionUtil.setLeftBorderColor(color.getIndex(), region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setBottomBorderColor(IndexedColors color) {
		RegionUtil.setBottomBorderColor(color.getIndex(), region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setRightBorderColor(IndexedColors color) {
		RegionUtil.setRightBorderColor(color.getIndex(), region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setTopBorderColor(IndexedColors color) {
		RegionUtil.setTopBorderColor(color.getIndex(), region, parent.sheet);
		return this;
	}

	public CellRangeBuilder setCellValue(Object value) {
		parent.createCell(region.getFirstRow(), region.getFirstColumn(), value);
		return this;
	}

	/**
	 * @deprecated use {@link CellRangeBuilder#createCellComment(String)} instead.
	 */
	@Deprecated
	public CellRangeBuilder setCellComment(String comment, String author, int row2, int col2) {
		parent.createCellComment(comment, author, region.getFirstRow(), region.getFirstColumn(), row2, col2);
		return this;
	}

	public PictureBuilder<CellRangeBuilder> createPicture(byte[] pictureData, int format) {
		int pictureIndex = parent.sheet.getWorkbook().addPicture(pictureData, format);
		return new PictureBuilder<>(this, parent.sheet, pictureIndex)
			.setRow1(region.getFirstRow())
			.setCol1(region.getFirstColumn())
			.setRow2(region.getLastRow() + 1)
			.setCol2(region.getLastColumn() + 1);
	}

	public CellCommentBuilder<CellRangeBuilder> createCellComment(String comment) {
		return new CellCommentBuilder<>(this, parent.sheet, comment)
			.setRow1(region.getFirstRow())
			.setCol1(region.getFirstColumn());
	}

	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.merge();
		return parent.addCellRange(firstRow, lastRow, firstCol, lastCol);
	}

	public SheetBuilder merge() {
		parent.sheet.addMergedRegion(region);
		return parent;
	}

	public SheetBuilder unsafeMerge() {
		parent.sheet.addMergedRegionUnsafe(region);
		return parent;
	}
}
