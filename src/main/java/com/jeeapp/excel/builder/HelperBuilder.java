package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author justice
 * @since 0.0.2
 */
public class HelperBuilder<B extends HelperBuilder<B>> extends CellStyleBuilder<B, SheetBuilder> {

	protected final Workbook workbook;

	protected final Sheet sheet;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final DataValidationHelper dataValidationHelper;

	public HelperBuilder(SheetBuilder parent, int row) {
		super(parent, row);
		this.workbook = parent.workbook;
		this.sheet = parent.sheet;
		this.drawing = parent.drawing;
		this.creationHelper = parent.creationHelper;
		this.dataValidationHelper = parent.dataValidationHelper;
	}

	public HelperBuilder(SheetBuilder parent, short column) {
		super(parent, column);
		this.workbook = parent.workbook;
		this.sheet = parent.sheet;
		this.drawing = parent.drawing;
		this.creationHelper = parent.creationHelper;
		this.dataValidationHelper = parent.dataValidationHelper;
	}

	public HelperBuilder(SheetBuilder parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.workbook = parent.workbook;
		this.sheet = parent.sheet;
		this.drawing = parent.drawing;
		this.creationHelper = parent.creationHelper;
		this.dataValidationHelper = parent.dataValidationHelper;
	}
}
