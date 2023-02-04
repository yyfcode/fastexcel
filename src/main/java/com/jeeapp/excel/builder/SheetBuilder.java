package com.jeeapp.excel.builder;

import java.util.Collection;
import java.util.Set;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.Assert;
import com.jeeapp.excel.model.Comment;
import com.jeeapp.excel.model.Row;

/**
 * @author justice
 */
@Slf4j
public class SheetBuilder extends SheetBuilderHelper<SheetBuilder> {

	private final WorkbookBuilder parent;

	protected final Sheet sheet;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final int maxRows;

	protected SheetBuilder(WorkbookBuilder parent, Sheet sheet) {
		super(parent, sheet);
		this.parent = parent;
		this.sheet = sheet;
		this.drawing = sheet.createDrawingPatriarch();
		this.creationHelper = sheet.getWorkbook().getCreationHelper();
		this.maxRows = sheet.getWorkbook().getSpreadsheetVersion().getMaxRows();
		initSheet(sheet);
	}

	/**
	 * 设置默认列宽
	 */
	@Override
	public SheetBuilder setDefaultColumnWidth(int width) {
		sheet.setDefaultColumnWidth(width);
		return this;
	}

	/**
	 * 设置默认行高
	 */
	@Override
	public SheetBuilder setDefaultRowHeight(int height) {
		sheet.setDefaultRowHeightInPoints(height);
		return this;
	}

	/**
	 * 设置列宽
	 */
	public SheetBuilder setColumnWidth(int column, int width) {
		sheet.setColumnWidth(column, width * 256);
		return this;
	}

	/**
	 * 创建多行
	 */
	public SheetBuilder createRows(Collection<? extends Row> rows) {
		for (Row row : rows) {
			createRow((Object[]) row.getCellValues());
			Set<Comment> comments = row.getComments();
			if (CollectionUtils.isNotEmpty(comments)) {
				for (Comment comment : comments) {
					matchingCell(sheet.getLastRowNum(), comment.getColNum())
						.createCellComment(comment.getText(), comment.getAuthor())
						.end();
				}
			}
		}
		return this;
	}

	/**
	 * 行构建器
	 */
	public <T> RowBuilderHelper<T> rowType(Class<T> type) {
		Assert.notNull(type, "Type must not be null");
		return new RowBuilderHelper<>(this, type);
	}

	/**
	 * 结束工作表
	 */
	public WorkbookBuilder end() {
		setSheetStyle(sheet);
		return parent;
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet() {
		return end().createSheet();
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet(String sheetName) {
		return end().createSheet(sheetName);
	}

	public Workbook build() {
		return end().build();
	}

	@Override
	protected SheetBuilder self() {
		return this;
	}
}
