package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.Assert;

/**
 * @author justice
 */
public class WorkbookBuilder extends CellBuilderHelper<WorkbookBuilder> {

	private final Workbook workbook;

	public WorkbookBuilder(Workbook workbook) {
		super(workbook);
		this.workbook = workbook;
	}

	/**
	 * 设置默认样式Excel工作簿
	 */
	public static WorkbookBuilder builder() {
		return new WorkbookBuilder(new SXSSFWorkbook())
			.setDefaultRowHeight(20)
			.matchingAll()
			.setFontHeight(12)
			.setFontName("微软雅黑")
			.setVerticalAlignment(VerticalAlignment.CENTER)
			.setAlignment(HorizontalAlignment.CENTER)
			.addCellStyle();
	}

	/**
	 * 自定义Excel工作簿
	 */
	public static WorkbookBuilder builder(Workbook workbook) {
		Assert.notNull(workbook, "workbook must not be null!");
		return new WorkbookBuilder(workbook);
	}

	/**
	 * 设置默认行高
	 */
	@Override
	public WorkbookBuilder setDefaultRowHeight(int height) {
		return super.setDefaultRowHeight(height);
	}

	/**
	 * 设置默认列宽
	 */
	@Override
	public WorkbookBuilder setDefaultColumnWidth(int width) {
		return super.setDefaultColumnWidth(width);
	}

	/**
	 * 打开工作表
	 */
	public SheetBuilder openSheet(String sheetName) {
		return new SheetBuilder(this, workbook.getSheet(WorkbookUtil.createSafeSheetName(sheetName)));
	}

	/**
	 * 打开工作表
	 */
	public SheetBuilder openSheet(int sheetIndex) {
		return new SheetBuilder(this, workbook.getSheetAt(sheetIndex));
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet() {
		return new SheetBuilder(this, workbook.createSheet());
	}

	/**
	 * 创建工作表
	 */
	public SheetBuilder createSheet(String sheetName) {
		return new SheetBuilder(this, workbook.createSheet(WorkbookUtil.createSafeSheetName(sheetName)));
	}

	public Workbook build() {
		return workbook;
	}

	@Override
	protected WorkbookBuilder self() {
		return this;
	}
}
