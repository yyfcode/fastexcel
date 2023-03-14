package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 * @author Justice
 * @since 0.0.2
 */
public abstract class CreationBuilder<B extends CreationBuilder<B>> extends CellStyleBuilder<B, SheetBuilder> {

	private final int firstRow;

	private final int lastRow;

	private final int firstCol;

	private final int lastCol;

	protected final Workbook workbook;

	protected final Sheet sheet;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final DataValidationHelper dataValidationHelper;

	protected CreationBuilder(SheetBuilder parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
		this.workbook = parent.workbook;
		this.sheet = parent.sheet;
		this.drawing = parent.drawing;
		this.creationHelper = parent.creationHelper;
		this.dataValidationHelper = parent.dataValidationHelper;
	}

	protected ValidationBuilder<B> createConstraint(int validationType, int operatorType, String firstFormula,
		String secondFormula, String[] explicitListValues, String dateFormat) {
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
			constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
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
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createExplicitListConstraint(String... explicitListValues) {
		DataValidationConstraint constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createFormulaListConstraint(String firstFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createTimeConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createTimeConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createDateConstraint(int operatorType, String firstFormula, String secondFormula, String dateFormat) {
		DataValidationConstraint constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createCustomConstraint(String firstFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createCustomConstraint(firstFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createIntegerConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createDecimalConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createTextLengthConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public B createPicture(byte[] pictureData, int format) {
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(firstRow);
		clientAnchor.setCol1(firstCol);
		clientAnchor.setRow2(lastRow + 1);
		clientAnchor.setCol2(lastCol + 1);
		int pictureIndex = workbook.addPicture(pictureData, format);
		drawing.createPicture(clientAnchor, pictureIndex);
		return self();
	}
}
