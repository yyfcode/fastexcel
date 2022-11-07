package com.jeeapp.excel.builder;

import java.util.Collection;
import java.util.Map;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author justice
 */
@Slf4j
public class SheetBuilder extends CellBuilder<SheetBuilder> {

	private final WorkbookBuilder parent;

	protected final Sheet sheet;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final int maxRows;

	protected int lastRow = -1;

	protected int lastCol = -1;

	public SheetBuilder(WorkbookBuilder parent, Sheet sheet) {
		super(parent);
		this.parent = parent;
		this.sheet = sheet;
		this.drawing = sheet.createDrawingPatriarch();
		this.creationHelper = sheet.getWorkbook().getCreationHelper();
		this.maxRows = sheet.getWorkbook().getSpreadsheetVersion().getMaxRows();
		super.initSheet(sheet);
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
	 * 设置默认列宽
	 */
	@Override
	public SheetBuilder autoSizeColumns(Integer... columns) {
		if (sheet instanceof SXSSFSheet) {
			((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		}
		for (Integer column : columns) {
			sheet.autoSizeColumn(column);
		}
		return this;
	}

	@Override
	protected void addColumnStyle(int column, Map<String, Object> properties) {
		super.addColumnStyle(column, properties);
		setColumnStyle(sheet, column);
	}

	/**
	 * 创建空行
	 */
	public SheetBuilder createRow() {
		lastRow = lastRow + 1;
		initRow(sheet.createRow(lastRow));
		lastCol = -1;
		return this;
	}

	/**
	 * 设置行高
	 */
	public SheetBuilder setRowHeight(int height) {
		Row row = sheet.getRow(lastRow);
		if (row != null) {
			row.setHeightInPoints(height);
		}
		return this;
	}

	/**
	 * 创建单行
	 */
	public SheetBuilder createRow(Object[] cells) {
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
			createRow(row.getCellValues());
			if (CollectionUtils.isNotEmpty(row.getComments())) {
				for (com.jeeapp.excel.model.Comment comment : row.getComments()) {
					createCellComment(comment.getText(),
						comment.getAuthor(),
						lastRow,
						comment.getColNum(),
						1,
						3
					);
				}
			}
		}
		return this;
	}

	/**
	 * 创建单元格
	 */
	private Cell createCell() {
		lastRow = lastRow == -1 ? 0 : lastRow;
		Row row = sheet.getRow(lastRow);
		if (row == null) {
			row = sheet.createRow(lastRow);
			initRow(row);
		}
		lastCol = lastCol + 1;
		Cell cell = row.createCell(lastCol);
		sheet.setActiveCell(new CellAddress(cell));
		return cell;
	}

	/**
	 * 创建有值的单元格(支持公式)
	 */
	public SheetBuilder createCell(Object value) {
		Cell cell = createCell();
		CellUtils.setCellValue(cell, value);
		super.setCellStyle(cell);
		return this;
	}

	/**
	 * 指定位置创建无值单元格
	 */
	private Cell createCell(int rowNum, int cellNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
			initRow(row);
		}
		Cell cell = row.getCell(cellNum);
		if (cell == null) {
			cell = row.createCell(cellNum);
		}
		sheet.setActiveCell(new CellAddress(cell));
		return cell;
	}

	/**
	 * 指定位置创建有值单元格
	 */
	public SheetBuilder createCell(int rowNum, int cellNum, Object value) {
		Cell cell = createCell(rowNum, cellNum);
		CellUtils.setCellValue(cell, value);
		super.setCellStyle(cell);
		return this;
	}

	/**
	 * 添加合并区域
	 */
	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.lastRow = Math.max(lastRow, this.lastRow);
		this.lastCol = Math.max(lastCol, this.lastCol);
		return new CellRangeBuilder(this, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * 创建验证
	 */
	public DataValidationBuilder createValidation() {
		CellAddress activeCell = sheet.getActiveCell();
		int row = activeCell.getRow();
		int column = activeCell.getColumn();
		return new DataValidationBuilder(this).setRegions(row, row, column, column);
	}

	/**
	 * 创建图片
	 */
	public PictureBuilder<SheetBuilder> createPicture(byte[] pictureData, int format) {
		CellAddress activeCell = sheet.getActiveCell();
		int row = activeCell.getRow();
		int column = activeCell.getColumn();
		int pictureIndex = sheet.getWorkbook().addPicture(pictureData, format);
		return new PictureBuilder<>(this, sheet, pictureIndex)
			.setRow1(row)
			.setCol1(column)
			.setSize(1, 1);
	}

	/**
	 * 创建批注
	 */
	public CellCommentBuilder<SheetBuilder> createCellComment(String comment) {
		CellAddress activeCell = sheet.getActiveCell();
		return new CellCommentBuilder<>(this, sheet, comment)
			.setRow1(activeCell.getRow())
			.setCol1(activeCell.getColumn());
	}

	/**
	 * 指定单元格添加批注
	 * @deprecated use {@link SheetBuilder#createCellComment(String)} instead.
	 */
	@Deprecated
	public SheetBuilder createCellComment(String comment, String author, int row1, int col1, int row2, int col2) {
		return new CellCommentBuilder<>(this, sheet, comment)
			.setRow1(row1)
			.setCol1(col1)
			.setSize(col2, row2)
			.setAuthor(author)
			.insert();
	}

	/**
	 * 当前单元格添加批注
	 * @deprecated use {@link SheetBuilder#createCellComment(String)} instead.
	 */
	@Deprecated
	public SheetBuilder createCellComment(String comment, String author, int row2, int col2) {
		CellAddress activeCell = sheet.getActiveCell();
		return new CellCommentBuilder<>(this, sheet, comment)
			.setRow1(activeCell.getRow())
			.setCol1(activeCell.getColumn())
			.setSize(col2, row2)
			.setAuthor(author)
			.insert();
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
		return parent.createSheet();
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet(String sheetName) {
		return parent.createSheet(sheetName);
	}

	/**
	 * 构建工作簿
	 */
	public Workbook build() {
		return parent.build();
	}
}
