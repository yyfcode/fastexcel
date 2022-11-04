package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

/**
 * @author justice
 */
public class CellRangeBuilder {

	private final SheetBuilder parent;

	private final Sheet sheet;

	private final CellRangeAddress region;

	public CellRangeBuilder(SheetBuilder parent, Sheet sheet, CellRangeAddress region) {
		this.parent = parent;
		this.sheet = sheet;
		this.region = region;
	}

	public CellRangeBuilder setBorder(BorderStyle border) {
		RegionUtil.setBorderLeft(border, region, sheet);
		RegionUtil.setBorderBottom(border, region, sheet);
		RegionUtil.setBorderRight(border, region, sheet);
		RegionUtil.setBorderTop(border, region, sheet);
		return this;
	}

	public CellRangeBuilder setBorderColor(int color) {
		RegionUtil.setLeftBorderColor(color, region, sheet);
		RegionUtil.setBottomBorderColor(color, region, sheet);
		RegionUtil.setRightBorderColor(color, region, sheet);
		RegionUtil.setTopBorderColor(color, region, sheet);
		return this;
	}

	public CellRangeBuilder setBorderLeft(BorderStyle border) {
		RegionUtil.setBorderLeft(border, region, sheet);
		return this;
	}

	public CellRangeBuilder setBorderBottom(BorderStyle border) {
		RegionUtil.setBorderBottom(border, region, sheet);
		return this;
	}

	public CellRangeBuilder setBorderRight(BorderStyle border) {
		RegionUtil.setBorderRight(border, region, sheet);
		return this;
	}

	public CellRangeBuilder setBorderTop(BorderStyle border) {
		RegionUtil.setBorderTop(border, region, sheet);
		return this;
	}

	public CellRangeBuilder setLeftBorderColor(int color) {
		RegionUtil.setLeftBorderColor(color, region, sheet);
		return this;
	}

	public CellRangeBuilder setBottomBorderColor(int color) {
		RegionUtil.setBottomBorderColor(color, region, sheet);
		return this;
	}

	public CellRangeBuilder setRightBorderColor(int color) {
		RegionUtil.setRightBorderColor(color, region, sheet);
		return this;
	}

	public CellRangeBuilder setTopBorderColor(int color) {
		RegionUtil.setTopBorderColor(color, region, sheet);
		return this;
	}

	public CellRangeBuilder setLeftBorderColor(IndexedColors color) {
		RegionUtil.setLeftBorderColor(color.getIndex(), region, sheet);
		return this;
	}

	public CellRangeBuilder setBottomBorderColor(IndexedColors color) {
		RegionUtil.setBottomBorderColor(color.getIndex(), region, sheet);
		return this;
	}

	public CellRangeBuilder setRightBorderColor(IndexedColors color) {
		RegionUtil.setRightBorderColor(color.getIndex(), region, sheet);
		return this;
	}

	public CellRangeBuilder setTopBorderColor(IndexedColors color) {
		RegionUtil.setTopBorderColor(color.getIndex(), region, sheet);
		return this;
	}

	public CellRangeBuilder setCellValue(Object value) {
		parent.createCell(region.getFirstRow(), region.getFirstColumn(), value);
		return this;
	}

	public CellRangeBuilder setCellComment(String comment, String author, int row2, int col2) {
		parent.createCellComment(comment, author, region.getFirstRow(), region.getFirstColumn(), row2, col2);
		return this;
	}

	public CellRangeBuilder setPictureData(byte[] pictureData, int format) {
		parent.createPicture(pictureData,
			format,
			region.getFirstRow(),
			region.getFirstColumn(),
			region.getLastRow() + 1,
			region.getLastColumn() + 1);
		return this;
	}

	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.merge();
		return parent.addCellRange(firstRow, lastRow, firstCol, lastCol);
	}

	public SheetBuilder merge() {
		if (region.getFirstRow() != region.getLastRow() || region.getFirstColumn() != region.getLastColumn()) {
			sheet.addMergedRegion(region);
		}
		return parent;
	}
}
