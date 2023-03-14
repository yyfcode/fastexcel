package com.jeeapp.excel.builder;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author Justice
 * @since 0.0.2
 */
@Slf4j
abstract class SheetBuilderHelper extends CellBuilderHelper<SheetBuilder> {

	protected final Sheet sheet;

	protected final int maxRows;

	protected final int maxColumns;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final DataValidationHelper dataValidationHelper;

	protected SheetBuilderHelper(CellBuilderHelper<?> parent, Sheet sheet) {
		super(parent);
		this.drawing = sheet.createDrawingPatriarch();
		this.creationHelper = workbook.getCreationHelper();
		this.dataValidationHelper = sheet.getDataValidationHelper();
		this.maxRows = workbook.getSpreadsheetVersion().getMaxRows();
		this.maxColumns = workbook.getSpreadsheetVersion().getMaxColumns();
		this.sheet = initSheet(sheet);
	}

	/**
	 * 创建空行
	 */
	public SheetBuilder createRow() {
		initRow(sheet.createRow(sheet.getLastRowNum() + 1));
		return self();
	}

	/**
	 * 创建单行
	 */
	public SheetBuilder createRow(Object... cells) {
		createRow();
		for (Object value : cells) {
			createCell(value);
		}
		return self();
	}

	/**
	 * 创建多行
	 */
	public SheetBuilder createRows(Object[][] rows) {
		for (Object[] cells : rows) {
			createRow(cells);
		}
		return self();
	}

	/**
	 * 创建有值的单元格(支持公式)
	 */
	public SheetBuilder createCell(Object value) {
		int lastRowNum = sheet.getLastRowNum() == -1 ? 0 : sheet.getLastRowNum();
		Row row = sheet.getRow(lastRowNum);
		if (row == null) {
			row = initRow(sheet.createRow(lastRowNum));
		}
		int lastCellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
		Cell cell = row.getCell(lastCellNum);
		if (cell == null) {
			cell = row.createCell(lastCellNum);
		}
		CellUtils.setCellValue(cell, value);
		setCellStyle(cell);
		return self();
	}

	/**
	 * 指定位置创建有值单元格
	 */
	public SheetBuilder createCell(CellAddress cellAddress, Object value) {
		Row row = sheet.getRow(cellAddress.getRow());
		if (row == null) {
			row = initRow(sheet.createRow(cellAddress.getRow()));
		}
		Cell cell = row.getCell(cellAddress.getColumn());
		if (cell == null) {
			cell = row.createCell(cellAddress.getColumn());
		}
		CellUtils.setCellValue(cell, value);
		setCellStyle(cell);
		return self();
	}

	/**
	 * 指定位置创建空单元格
	 */
	public SheetBuilder createCell(CellAddress cellAddress) {
		return createCell(cellAddress, null);
	}

	/**
	 * 指定位置创建有值单元格
	 */
	public SheetBuilder createCell(int row, int column, Object value) {
		return createCell(new CellAddress(row, column), value);
	}

	/**
	 * 指定位置创建空单元格
	 */
	public SheetBuilder createCell(int row, int column) {
		return createCell(row, column, null);
	}

	/**
	 * 匹配行
	 */
	@Override
	public RowBuilder matchingRow(int row) {
		return new RowBuilder(self(), row);
	}

	/**
	 * 匹配最后一行
	 */
	public RowBuilder matchingLastRow() {
		return new RowBuilder(self(), sheet.getLastRowNum());
	}

	/**
	 * 匹配列
	 */
	@Override
	public ColumnBuilder matchingColumn(int column) {
		return new ColumnBuilder(self(), column);
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder matchingCell(CellAddress cellAddress) {
		return new CellBuilder(self(), cellAddress);
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder matchingCell(int row, int column) {
		return matchingCell(new CellAddress(row, column));
	}

	/**
	 * 匹配活动单元格的位置
	 */
	public CellBuilder matchingActiveCell() {
		return matchingCell(sheet.getActiveCell());
	}

	/**
	 * 匹配最后一个单元格
	 */
	public CellBuilder matchingLastCell() {
		int lastRowNum = sheet.getLastRowNum() == -1 ? 0 : sheet.getLastRowNum();
		Row row = sheet.getRow(lastRowNum);
		if (row == null) {
			row = initRow(sheet.createRow(lastRowNum));
		}
		short lastCellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
		Cell cell = row.getCell(lastCellNum);
		if (cell == null) {
			cell = row.createCell(lastCellNum);
		}
		return matchingCell(new CellAddress(cell));
	}

	/**
	 * 匹配最后一行上的单元格
	 */
	public CellBuilder matchingLastRowCell(int column) {
		return matchingCell(new CellAddress(sheet.getLastRowNum(), column));
	}

	/**
	 * 设置默认列宽
	 */
	@Override
	public SheetBuilder setDefaultColumnWidth(int width) {
		sheet.setDefaultColumnWidth(width);
		return self();
	}

	/**
	 * 设置默认行高
	 */
	@Override
	public SheetBuilder setDefaultRowHeight(int height) {
		sheet.setDefaultRowHeightInPoints(height);
		return self();
	}

	/**
	 * 自动换行
	 */
	public SheetBuilder setAutoBreaks(Boolean autoBreaks) {
		sheet.setAutobreaks(autoBreaks);
		return self();
	}

	/**
	 * 匹配区域
	 * @param firstRow 起始行
	 * @param lastRow 结束行必须等于或大于 {@code firstRow}
	 * @param firstCol 起始列
	 * @param lastCol 结束列必须等于或大于 {@code firstCol}
	 */
	@Override
	public CellRangeBuilder matchingRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
		return new CellRangeBuilder(self(), firstRow, lastRow, firstCol, lastCol);

	}

	/**
	 * 指定单元格添加批注
	 * @deprecated use {@link SheetBuilderHelper#matchingCell(CellAddress)} instead.
	 */
	@Deprecated
	public SheetBuilder createCellComment(String comment, String author, int row1, int col1, int row2, int col2) {
		return matchingCell(new CellAddress(row1, col1))
			.createCellComment(comment, author, row2, col2)
			.addCellStyle();
	}

	/**
	 * 当前单元格添加批注
	 * @deprecated use {@link SheetBuilderHelper#matchingLastCell()} instead.
	 */
	@Deprecated
	public SheetBuilder createCellComment(String comment, String author, int row2, int col2) {
		return matchingLastCell()
			.createCellComment(comment, author, row2, col2)
			.addCellStyle();
	}
}
