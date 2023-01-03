package com.jeeapp.excel.builder;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author Justice
 * @since 0.0.2
 */
@Slf4j
abstract class RowBuilderHelper<B extends RowBuilderHelper<B>> extends CellBuilderHelper<B> {

	protected final Sheet sheet;

	protected RowBuilderHelper(CellBuilderHelper<?> parent, Sheet sheet) {
		super(parent);
		this.sheet = sheet;
	}

	/**
	 * 设置行高
	 */
	public B setRowHeight(int height) {
		Row row = sheet.getRow(sheet.getLastRowNum());
		if (row != null) {
			row.setHeightInPoints(height);
		}
		return self();
	}

	/**
	 * 创建空行
	 */
	public B createRow() {
		Row row = sheet.createRow(sheet.getLastRowNum() + 1);
		setRowStyle(sheet, row);
		return self();
	}

	/**
	 * 创建单行
	 */
	public B createRow(Object... cells) {
		createRow();
		for (Object value : cells) {
			createCell(value);
		}
		return self();
	}

	/**
	 * 创建多行
	 */
	public B createRows(Object[][] rows) {
		for (Object[] cells : rows) {
			createRow(cells);
		}
		return self();
	}

	/**
	 * 创建单元格
	 */
	protected Cell createCell() {
		int lastRowNum = sheet.getLastRowNum() == -1 ? 0 : sheet.getLastRowNum();
		Row row = sheet.getRow(lastRowNum);
		if (row == null) {
			row = sheet.createRow(lastRowNum);
			setRowStyle(sheet, row);
		}
		int lastCellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
		Cell cell = row.getCell(lastCellNum);
		if (cell == null) {
			cell = row.createCell(lastCellNum);
		}
		return cell;
	}

	/**
	 * 创建有值的单元格(支持公式)
	 */
	public B createCell(Object value) {
		Cell cell = createCell();
		CellUtils.setCellValue(cell, value);
		setCellStyle(cell);
		return self();
	}

	/**
	 * 指定位置创建无值单元格
	 */
	protected Cell createCell(CellAddress cellAddress) {
		Row row = sheet.getRow(cellAddress.getRow());
		if (row == null) {
			row = sheet.createRow(cellAddress.getRow());
			setRowStyle(sheet, row);
		}
		Cell cell = row.getCell(cellAddress.getColumn());
		if (cell == null) {
			cell = row.createCell(cellAddress.getColumn());
		}
		return cell;
	}

	/**
	 * 指定位置创建有值单元格
	 */
	public B createCell(CellAddress cellAddress, Object value) {
		Cell cell = createCell(cellAddress);
		CellUtils.setCellValue(cell, value);
		setCellStyle(cell);
		return self();
	}

	/**
	 * 指定位置创建有值单元格
	 */
	public B createCell(int row, int column, Object value) {
		return createCell(new CellAddress(row, column), value);
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder<B> matchingCell(CellAddress cellAddress) {
		return new CellBuilder<>(self(), cellAddress);
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder<B> matchingCell(int row, int column) {
		return matchingCell(new CellAddress(row, column));
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder<B> matchingCell() {
		return matchingCell(sheet.getLastRowNum(), sheet.getRow(sheet.getLastRowNum()).getLastCellNum());
	}

	/**
	 * 匹配区域
	 * @param firstRow 起始行
	 * @param lastRow 结束行必须等于或大于 {@code firstRow}
	 * @param firstCol 起始列
	 * @param lastCol 结束列必须等于或大于 {@code firstCol}
	 */
	@Override
	public CellRangeBuilder<B> matchingRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
		return new CellRangeBuilder<>(self(), firstRow, lastRow, firstCol, lastCol);

	}

	/**
	 * 指定单元格添加批注
	 * @deprecated use {@link RowBuilderHelper#matchingCell(CellAddress)} instead.
	 */
	@Deprecated
	public B createCellComment(String comment, String author, int row1, int col1, int row2, int col2) {
		return matchingCell(new CellAddress(row1, col1))
			.setCommentText(comment)
			.setCommentSize(row2, col2)
			.setCommentAuthor(author)
			.end();
	}

	/**
	 * 当前单元格添加批注
	 * @deprecated use {@link RowBuilderHelper#matchingCell()} instead.
	 */
	@Deprecated
	public B createCellComment(String comment, String author, int row2, int col2) {
		return matchingCell()
			.setCommentText(comment)
			.setCommentSize(row2, col2)
			.setCommentAuthor(author)
			.end();
	}
}
