package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 * @author Justice
 * @since 0.0.2
 */
public class DataValidationBuilder<B extends DataValidationBuilder<B, P>, P extends SheetBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	private final CellRangeAddressList regions;

	protected DataValidationBuilder(P parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.parent = parent;
		CellRangeAddressList regions = new CellRangeAddressList();
		regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		this.regions = regions;
	}

	protected ValidationBuilder<B, P> createConstraint(int validationType, int operatorType, String firstFormula,
		String secondFormula, String[] explicitListValues, String dateFormat) {
		DataValidationConstraint constraint = null;
		if (validationType == ValidationType.LIST) {
			if (explicitListValues != null) {
				constraint = parent.dataValidationHelper.createExplicitListConstraint(explicitListValues);
			} else {
				constraint = parent.dataValidationHelper.createFormulaListConstraint(firstFormula);
			}
		}
		if (validationType == ValidationType.TIME) {
			constraint = parent.dataValidationHelper.createTimeConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DATE) {
			constraint = parent.dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		}
		if (validationType == ValidationType.FORMULA) {
			constraint = parent.dataValidationHelper.createCustomConstraint(firstFormula);
		}
		if (validationType == ValidationType.INTEGER) {
			constraint = parent.dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DECIMAL) {
			constraint = parent.dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.TEXT_LENGTH) {
			constraint = parent.dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		}
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createExplicitListConstraint(String... explicitListValues) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createExplicitListConstraint(explicitListValues);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createFormulaListConstraint(String firstFormula) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createFormulaListConstraint(firstFormula);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createTimeConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createTimeConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createDateConstraint(int operatorType, String firstFormula, String secondFormula, String dateFormat) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createCustomConstraint(String firstFormula) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createCustomConstraint(firstFormula);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createIntegerConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createDecimalConstraint(int operatorType, String firstFormula, String secondFormula) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	public ValidationBuilder<B, P> createTextLengthConstraint(int operatorType, String firstFormula,
		String secondFormula) {
		DataValidationConstraint constraint = parent.dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		DataValidation validation = parent.dataValidationHelper.createValidation(constraint, regions);
		return new ValidationBuilder<>(self(), validation);
	}

	protected B addValidationData(DataValidation validation) {
		parent.sheet.addValidationData(validation);
		return self();
	}

	@Override
	@SuppressWarnings("unchecked")
	protected B self() {
		return (B) this;
	}

	public static class ValidationBuilder<B extends DataValidationBuilder<B, P>, P extends SheetBuilderHelper<P>> {

		private final DataValidationBuilder<B, P> parent;

		private final DataValidation validation;

		private boolean allowedEmptyCell = true;

		private boolean showPromptBox = false;

		private boolean showErrorBox = false;

		private boolean suppress = true;

		private int errorStyle = ErrorStyle.WARNING;

		private String errorBoxTitle;

		private String errorBoxText;

		private String promptBoxTitle;

		private String promptBoxText;

		protected ValidationBuilder(DataValidationBuilder<B, P> parent, DataValidation validation) {
			this.parent = parent;
			this.validation = validation;
		}

		public ValidationBuilder<B, P> allowedEmptyCell(boolean allowedEmptyCell) {
			this.allowedEmptyCell = allowedEmptyCell;
			return this;
		}

		public ValidationBuilder<B, P> setErrorStyle(int errorStyle) {
			this.errorStyle = errorStyle;
			return this;
		}

		public ValidationBuilder<B, P> setSuppressDropDownArrow(boolean suppress) {
			this.suppress = suppress;
			return this;
		}

		public ValidationBuilder<B, P> showErrorBox(String errorBoxTitle, String errorBoxText) {
			showErrorBox(true, errorBoxTitle, errorBoxText);
			return this;
		}

		public ValidationBuilder<B, P> showPromptBox(String promptBoxTitle, String promptBoxText) {
			showPromptBox(true, promptBoxTitle, promptBoxText);
			return this;
		}

		protected ValidationBuilder<B, P> showErrorBox(boolean showErrorBox, String errorBoxTitle, String errorBoxText) {
			this.showErrorBox = showErrorBox;
			this.errorBoxTitle = errorBoxTitle;
			this.errorBoxText = errorBoxText;
			return this;
		}

		protected ValidationBuilder<B, P> showPromptBox(boolean showPromptBox, String promptBoxTitle, String promptBoxText) {
			this.showPromptBox = showPromptBox;
			this.promptBoxTitle = promptBoxTitle;
			this.promptBoxText = promptBoxText;
			return this;
		}

		public DataValidationBuilder<B, P> addValidationData() {
			validation.setEmptyCellAllowed(allowedEmptyCell);
			validation.setSuppressDropDownArrow(suppress);
			validation.setErrorStyle(errorStyle);
			validation.setShowErrorBox(showErrorBox);
			validation.setShowPromptBox(showPromptBox);
			validation.createErrorBox(errorBoxTitle, errorBoxText);
			validation.createPromptBox(promptBoxTitle, promptBoxText);
			return parent.addValidationData(validation);
		}
	}
}
