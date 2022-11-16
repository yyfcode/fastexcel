package com.jeeapp.excel.builder;

import java.util.Collection;
import java.util.Set;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import com.jeeapp.excel.model.Comment;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author justice
 */
@Slf4j
public class SheetBuilder extends CellBuilderHelper<SheetBuilder> {

	private final WorkbookBuilder parent;

	protected final Sheet sheet;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final int maxRows;

	public SheetBuilder(WorkbookBuilder parent, Sheet sheet) {
		super(parent);
		this.parent = parent;
		this.sheet = sheet;
		this.drawing = sheet.createDrawingPatriarch();
		this.creationHelper = sheet.getWorkbook().getCreationHelper();
		this.maxRows = sheet.getWorkbook().getSpreadsheetVersion().getMaxRows();
		initSheet(sheet);
	}

	/**
	 * 工作表列宽
	 */
	@Override
	public SheetBuilder setDefaultColumnWidth(int width) {
		sheet.setDefaultColumnWidth(width);
		return this;
	}

	/**
	 * 设置列宽
	 */
	@Override
	public SheetBuilder setColumnWidth(int column, int width) {
		sheet.setColumnWidth(column, width * 256);
		return this;
	}

	/**
	 * 创建空行
	 */
	public SheetBuilder createRow() {
		sheet.createRow(sheet.getLastRowNum() + 1);
		return this;
	}

	/**
	 * 设置行高
	 */
	public SheetBuilder setRowHeight(int height) {
		Row row = sheet.getRow(sheet.getLastRowNum());
		if (row != null) {
			row.setHeightInPoints(height);
		}
		return this;
	}

	/**
	 * 创建单行
	 */
	public SheetBuilder createRow(Object... cells) {
		createRow();
		for (Object value : cells) {
			createCell(value);
		}
		return this;
	}

	/**
	 * 创建多行
	 */
	public SheetBuilder createRows(Object[][] rows) {
		for (Object[] cells : rows) {
			createRow(cells);
		}
		return this;
	}

	/**
	 * 创建多行
	 */
	public SheetBuilder createRows(Collection<? extends com.jeeapp.excel.model.Row> rows) {
		for (com.jeeapp.excel.model.Row row : rows) {
			createRow((Object[]) row.getCellValues());
			Set<Comment> comments = row.getComments();
			if (CollectionUtils.isNotEmpty(comments)) {
				for (com.jeeapp.excel.model.Comment comment : comments) {
					matchingCell(sheet.getLastRowNum(), comment.getColNum())
						.setCommentSize(1, 3)
						.setCommentText(comment.getText())
						.setCommentAuthor(comment.getAuthor())
						.end();
				}
			}
		}
		return this;
	}

	/**
	 * 创建单元格
	 */
	protected Cell createCell() {
		int lastRowNum = sheet.getLastRowNum() == -1 ? 0 : sheet.getLastRowNum();
		Row row = sheet.getRow(lastRowNum);
		if (row == null) {
			row = sheet.createRow(lastRowNum);
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
	public SheetBuilder createCell(Object value) {
		Cell cell = createCell();
		CellUtils.setCellValue(cell, value);
		return this;
	}

	/**
	 * 指定位置创建无值单元格
	 */
	protected Cell createCell(CellAddress cellAddress) {
		Row row = sheet.getRow(cellAddress.getRow());
		if (row == null) {
			row = sheet.createRow(cellAddress.getRow());
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
	public SheetBuilder createCell(CellAddress cellAddress, Object value) {
		Cell cell = createCell(cellAddress);
		CellUtils.setCellValue(cell, value);
		return this;
	}

	/**
	 * 指定位置创建有值单元格
	 */
	public SheetBuilder createCell(int row, int column, Object value) {
		return createCell(new CellAddress(row, column), value);
	}


	/**
	 * @deprecated use {@link SheetBuilder#matchingRegion(int, int, int, int)} instead.
	 */
	@Deprecated
	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		return new CellRangeBuilder(this, firstRow, lastRow, firstCol, lastCol);
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
		return new CellRangeBuilder(this, firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder matchingCell(CellAddress cellAddress) {
		return new CellBuilder(this, cellAddress);
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder matchingCell(int row, int column) {
		return matchingCell(new CellAddress(row, column));
	}

	/**
	 * 匹配单元格
	 */
	public CellBuilder matchingCell() {
		return matchingCell(sheet.getLastRowNum(), sheet.getRow(sheet.getLastRowNum()).getLastCellNum());
	}

	/**
	 * 指定单元格添加批注
	 * @deprecated use {@link SheetBuilder#matchingCell(CellAddress)} instead.
	 */
	@Deprecated
	public SheetBuilder createCellComment(String comment, String author, int row1, int col1, int row2, int col2) {
		return matchingCell(new CellAddress(row1, col1))
			.setCommentText(comment)
			.setCommentSize(row2, col2)
			.setCommentAuthor(author)
			.end();
	}

	/**
	 * 当前单元格添加批注
	 * @deprecated use {@link SheetBuilder#matchingCell()} instead.
	 */
	@Deprecated
	public SheetBuilder createCellComment(String comment, String author, int row2, int col2) {
		return matchingCell()
			.setCommentText(comment)
			.setCommentSize(row2, col2)
			.setCommentAuthor(author)
			.end();
	}

	/**
	 * 行构建器
	 */
	public <T> RowBuilder<T> rowType(Class<T> type) {
		return new RowBuilder<>(this, type);
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet() {
		return end().parent.createSheet();
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet(String sheetName) {
		return end().parent.createSheet(sheetName);
	}

	@Override
	public Workbook build() {
		return end().parent.build();
	}

	protected SheetBuilder end() {
		return end(sheet);
	}

	@Override
	protected SheetBuilder self() {
		return this;
	}
}
