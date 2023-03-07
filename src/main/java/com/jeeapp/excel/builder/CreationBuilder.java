package com.jeeapp.excel.builder;

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
public abstract class CreationBuilder<B extends CreationBuilder<B, P>, P extends SheetBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	private final CellRangeAddressList regions;

	protected final Workbook workbook;

	protected final Sheet sheet;

	protected final Drawing<?> drawing;

	protected final CreationHelper creationHelper;

	protected final DataValidationHelper dataValidationHelper;

	protected CreationBuilder(P parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.parent = parent;
		this.workbook = parent.workbook;
		this.sheet = parent.sheet;
		this.drawing = parent.drawing;
		this.creationHelper = parent.creationHelper;
		this.dataValidationHelper = parent.dataValidationHelper;
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		this.regions = regions;
	}

	protected ValidationBuilder<B, P> createConstraint(int validationType, int operatorType, String firstFormula,
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
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createExplicitListConstraint(String... explicitListValues) {
		DataValidationConstraint constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createFormulaListConstraint(String firstFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createTimeConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createTimeConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createDateConstraint(int operatorType, String firstFormula, String secondFormula, String dateFormat) {
		DataValidationConstraint constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createCustomConstraint(String firstFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createCustomConstraint(firstFormula);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createIntegerConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createDecimalConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createTextLengthConstraint(int operatorType, String firstFormula,
		String secondFormula) {
		DataValidationConstraint constraint = dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	@Override
	@SuppressWarnings("unchecked")
	protected B self() {
		return (B) this;
	}
}
