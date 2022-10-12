package io.github.yyfcode.fastexcel.builder;

import java.util.Arrays;
import java.util.Collection;
import java.util.Map;

import io.github.yyfcode.fastexcel.util.CellUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author justice
 */
@Slf4j
public class SheetBuilder extends CellBuilderHelper<SheetBuilder> {

	private final WorkbookBuilder parent;

	private final Sheet sheet;

	private final Drawing<?> drawing;

	private final CreationHelper creationHelper;

	private final DataValidationHelper dataValidationHelper;

	private final int maxRows;

	private int lastRow = -1;

	private int lastCol = -1;

	public SheetBuilder(WorkbookBuilder parent, Sheet sheet) {
		super(parent);
		this.parent = parent;
		this.sheet = sheet;
		this.drawing = sheet.createDrawingPatriarch();
		this.creationHelper = sheet.getWorkbook().getCreationHelper();
		this.dataValidationHelper = sheet.getDataValidationHelper();
		this.maxRows = sheet.getWorkbook().getSpreadsheetVersion().getMaxRows();
		super.initSheetStyle(sheet);
	}

	/**
	 * 工作表行高
	 */
	@Override
	public SheetBuilder setDefaultRowHeight(int height) {
		sheet.setDefaultRowHeight((short) (height * 20));
		return this;
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
	 * 添加合并区域
	 */
	public CellRangeBuilder addCellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.lastRow = Math.max(lastRow, this.lastRow);
		this.lastCol = Math.max(lastCol, this.lastCol);
		return new CellRangeBuilder(this, sheet, new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * 创建空行
	 */
	public SheetBuilder createRow() {
		lastRow = lastRow + 1;
		sheet.createRow(lastRow);
		lastCol = -1;
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
	public SheetBuilder createRows(Collection<? extends Row> rows) {
		for (Row row : rows) {
			createRow(row.getCellValues());
			if (CollectionUtils.isNotEmpty(row.getComments())) {
				for (Comment comment : row.getComments()) {
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
	 * 创建单元格(支持公式)
	 */
	public SheetBuilder createCell(Object value) {
		lastRow = lastRow == -1 ? 0 : lastRow;
		Row row = sheet.getRow(lastRow);
		if (row == null) {
			row = sheet.createRow(lastRow);
		}
		lastCol = lastCol == -1 ? 0 : lastCol + 1;
		Cell cell = row.createCell(lastCol);
		CellUtils.setCellValue(cell, value);
		super.setCellStyle(cell);
		return this;
	}

	/**
	 * 指定位置创建单元格
	 */
	public SheetBuilder createCell(int row1, int col1, Object value) {
		Row row = sheet.getRow(row1);
		if (row == null) {
			row = sheet.createRow(row1);
		}
		Cell cell = row.getCell(col1);
		if (cell == null) {
			cell = row.createCell(col1);
		}
		if (value != null) {
			CellUtils.setCellValue(cell, value);
		}
		super.setCellStyle(cell);
		return this;
	}

	/**
	 * 指定位置添加批注
	 */
	public SheetBuilder createCellComment(String comment, String author, int row1, int col1, int row2, int col2) {
		Row row = sheet.getRow(row1);
		if (row == null) {
			row = sheet.createRow(row1);
		}
		Cell cell = row.getCell(col1);
		if (cell == null) {
			cell = row.createCell(col1);
		}
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setCol1(col1);
		clientAnchor.setCol2(col1 + col2);
		clientAnchor.setRow1(row1);
		clientAnchor.setRow2(row1 + row2);
		clientAnchor.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE);
		Comment cellComment = drawing.createCellComment(clientAnchor);
		cellComment.setString(creationHelper.createRichTextString(comment));
		cellComment.setAuthor(author);
		cell.setCellComment(cellComment);
		return this;
	}

	/**
	 * 当前单元格添加批注
	 */
	public SheetBuilder createCellComment(String comment, String author, int row2, int col2) {
		lastRow = lastRow == -1 ? 0 : lastRow;
		Row row = sheet.getRow(lastRow);
		if (row == null) {
			row = sheet.createRow(lastRow);
		}
		lastCol = lastCol == -1 ? 0 : lastCol;
		Cell cell = row.getCell(lastCol);
		if (cell == null) {
			cell = row.createCell(lastCol);
		}
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setCol1(cell.getColumnIndex());
		clientAnchor.setCol2(cell.getColumnIndex() + row2);
		clientAnchor.setRow1(cell.getRowIndex());
		clientAnchor.setRow2(cell.getRowIndex() + col2);
		clientAnchor.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE);
		Comment cellComment = drawing.createCellComment(clientAnchor);
		cellComment.setString(creationHelper.createRichTextString(comment));
		cellComment.setAuthor(author);
		cell.setCellComment(cellComment);
		return this;
	}

	/**
	 * 给列添加数据验证
	 */
	protected SheetBuilder addValidationData(CellRangeAddressList cellRangeAddressList,
		int validationType, int operatorType, String firstFormula, String secondFormula, String[] explicitListValues,
		boolean allowEmpty, int errorStyle, boolean showPromptBox, String promptBoxTitle, String promptBoxText,
		boolean showErrorBox, String errorBoxTitle, String errorBoxText) {
		DataValidationConstraint constraint = null;
		if (validationType == ValidationType.LIST) {
			if (explicitListValues != null) {
				constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
			} else {
				constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
			}
		}
		if (validationType == ValidationType.TIME) {
			constraint = dataValidationHelper.createTimeConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DATE) {
			constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, null);
		}
		if (validationType == ValidationType.FORMULA) {
			constraint = dataValidationHelper.createCustomConstraint(firstFormula);
		}
		if (validationType == ValidationType.INTEGER) {
			constraint = dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DECIMAL) {
			constraint = dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.TEXT_LENGTH) {
			constraint = dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		}
		if (constraint != null) {
			DataValidation validation = dataValidationHelper.createValidation(constraint, cellRangeAddressList);
			validation.setEmptyCellAllowed(allowEmpty);
			validation.setErrorStyle(errorStyle);
			if (showErrorBox) {
				validation.setShowErrorBox(true);
				if (StringUtils.isBlank(errorBoxText)) {
					errorBoxText = createDefaultErrorBoxText(validationType, operatorType, firstFormula,
						secondFormula, explicitListValues);
				}
				validation.createErrorBox(errorBoxTitle, errorBoxText);
			}
			if (showPromptBox) {
				validation.setShowPromptBox(true);
				validation.createPromptBox(promptBoxTitle, promptBoxText);
			}
			sheet.addValidationData(validation);
		}
		return this;
	}

	/**
	 * 默认的错误提示
	 */
	protected String createDefaultErrorBoxText(int validationType, int operatorType, String firstFormula,
		String secondFormula, String[] explicitListValues) {
		String type = "";
		if (validationType == ValidationType.LIST) {
			if (explicitListValues != null) {
				return String.format("必须是%s其中之一", Arrays.toString(explicitListValues));
			} else {
				return String.format("数据有误，验证规则：%s!", firstFormula);
			}
		}
		if (validationType == ValidationType.TIME) {
			type = "时间";
		}
		if (validationType == ValidationType.DATE) {
			type = "日期";
		}
		if (validationType == ValidationType.FORMULA) {
			return String.format("数据有误，验证规则：%s!", firstFormula);
		}
		if (validationType == ValidationType.TEXT_LENGTH) {
			type = "长度";
		}
		if (operatorType == OperatorType.BETWEEN) {
			return String.format("%s必须在%s和%s之间!", type, firstFormula, secondFormula);
		}
		if (operatorType == OperatorType.NOT_BETWEEN) {
			return String.format("%s不能在%s和%s之间!", type, firstFormula, secondFormula);
		}
		if (operatorType == OperatorType.EQUAL) {
			return String.format("%s必须等于%s!", type, firstFormula);
		}
		if (operatorType == OperatorType.NOT_EQUAL) {
			return String.format("%s不能等于%s!", type, firstFormula);
		}
		if (operatorType == OperatorType.GREATER_THAN) {
			return String.format("%s必须大于%s!", type, firstFormula);
		}
		if (operatorType == OperatorType.LESS_THAN) {
			return String.format("%s必须小于%s!", type, firstFormula);
		}
		if (operatorType == OperatorType.GREATER_OR_EQUAL) {
			return String.format("%s必须大于或等于%s!", type, firstFormula);
		}
		if (operatorType == OperatorType.LESS_OR_EQUAL) {
			return String.format("%s必须小于或等于%s!", type, firstFormula);
		}
		return "数据验证不通过!";
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

	/**
	 * 获取最后一行
	 */
	protected int getLastRow() {
		return lastRow;
	}

	protected int getMaxRows() {
		return this.maxRows;
	}
}
