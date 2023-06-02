package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 * @author Justice
 * @since 0.0.2
 */
abstract class CreationBuilder<B extends CreationBuilder<B>> extends CellStyleBuilder<B, SheetBuilder> {

	protected final SheetBuilder parent;

	private final int firstRow;

	private final int lastRow;

	private final int firstCol;

	private final int lastCol;

	protected CreationBuilder(SheetBuilder parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.parent = parent;
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}

	@Deprecated
	protected ValidationBuilder<B> createConstraint(int validationType, int operatorType, String firstFormula,
		String secondFormula, String[] explicitListValues, String dateFormat) {
		DataValidationConstraint constraint = null;
		if (validationType == ValidationType.LIST) {
			if (explicitListValues != null) {
				constraint = parent.helper.getDataValidationHelper().createExplicitListConstraint(explicitListValues);
			} else {
				constraint = parent.helper.getDataValidationHelper().createFormulaListConstraint(firstFormula);
			}
		}
		if (validationType == ValidationType.TIME) {
			constraint = parent.helper.getDataValidationHelper().createTimeConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DATE) {
			constraint = parent.helper.getDataValidationHelper().createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		}
		if (validationType == ValidationType.FORMULA) {
			constraint = parent.helper.getDataValidationHelper().createCustomConstraint(firstFormula);
		}
		if (validationType == ValidationType.INTEGER) {
			constraint = parent.helper.getDataValidationHelper().createIntegerConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DECIMAL) {
			constraint = parent.helper.getDataValidationHelper().createDecimalConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.TEXT_LENGTH) {
			constraint = parent.helper.getDataValidationHelper().createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		}
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createExplicitListConstraint(String... explicitListValues) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createExplicitListConstraint(explicitListValues);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createFormulaListConstraint(String firstFormula) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createFormulaListConstraint(firstFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createTimeConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createTimeConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createDateConstraint(int operatorType, String firstFormula, String secondFormula, String dateFormat) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createCustomConstraint(String firstFormula) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createCustomConstraint(firstFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createIntegerConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createIntegerConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createDecimalConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createDecimalConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B> createTextLengthConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.helper.getDataValidationHelper().createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		DataValidation validation = parent.helper.getDataValidationHelper().createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public B createPicture(byte[] pictureData, int format) {
		ClientAnchor clientAnchor = parent.helper.createClientAnchor();
		clientAnchor.setRow1(firstRow);
		clientAnchor.setCol1(firstCol);
		clientAnchor.setRow2(lastRow + 1);
		clientAnchor.setCol2(lastCol + 1);
		int pictureIndex = parent.helper.addPicture(pictureData, format);
		parent.helper.createPicture(clientAnchor, pictureIndex);
		return self();
	}
}
