package com.jeeapp.excel.builder;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.Assert;

/**
 * @author justice
 */
@Slf4j
public class SheetBuilder extends SheetBuilderHelper<SheetBuilder> {

	private final WorkbookBuilder parent;

	protected SheetBuilder(WorkbookBuilder parent, Sheet sheet) {
		super(parent, sheet);
		this.parent = parent;
	}

	public <T> TableBuilder<T> rowType(Class<T> type) {
		Assert.notNull(type, "Type must not be null");
		return new TableBuilder<>(this, type);
	}

	public WorkbookBuilder end() {
		return parent.setSheetStyle(sheet);
	}

	public SheetBuilder createSheet() {
		return end().createSheet();
	}

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
