package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 * @author Justice
 * @since 0.0.2
 */
@SuppressWarnings("unchecked")
public class DataValidationBuilder<B extends DataValidationBuilder<B, P>, P extends RowBuilderHelper<P>> extends CellStyleBuilder<B, P> {

	private final P parent;

	protected final CreationHelper creationHelper;

	private final DataValidationHelper dataValidationHelper;

	private DataValidationConstraint constraint;

	private final int firstRow;

	private final int lastRow;

	private final int firstCol;

	private final int lastCol;

	private boolean allowedEmptyCell = true;

	private boolean showPromptBox = false;

	private boolean showErrorBox = false;

	private boolean suppress = true;

	private int errorStyle = ErrorStyle.WARNING;

	private String errorBoxTitle;

	private String errorBoxText;

	private String promptBoxTitle;

	private String promptBoxText;

	protected DataValidationBuilder(P parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.parent = parent;
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
		this.dataValidationHelper = parent.sheet.getDataValidationHelper();
		this.creationHelper = parent.sheet.getWorkbook().getCreationHelper();
	}

	protected B createConstraint(int validationType, int operatorType, String firstFormula,
		String secondFormula, String[] explicitListValues, String dateFormat) {
		if (validationType == ValidationType.LIST) {
			if (explicitListValues != null) {
				this.constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
			} else {
				this.constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
			}
		}
		if (validationType == ValidationType.TIME) {
			this.constraint = dataValidationHelper.createTimeConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DATE) {
			this.constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		}
		if (validationType == ValidationType.FORMULA) {
			this.constraint = dataValidationHelper.createCustomConstraint(firstFormula);
		}
		if (validationType == ValidationType.INTEGER) {
			this.constraint = dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.DECIMAL) {
			this.constraint = dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		}
		if (validationType == ValidationType.TEXT_LENGTH) {
			this.constraint = dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		}
		return self();
	}

	public B createExplicitListConstraint(String... explicitListValues) {
		constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
		return self();
	}

	public B createFormulaListConstraint(String firstFormula) {
		constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
		return self();
	}

	public B createTimeConstraint(String firstFormula) {
		constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
		return self();
	}

	public B createDateConstraint(int operatorType, String firstFormula, String secondFormula, String dateFormat) {
		constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		return self();
	}

	public B createCustomConstraint(String firstFormula) {
		constraint = dataValidationHelper.createCustomConstraint(firstFormula);
		return self();
	}

	public B createIntegerConstraint(int operatorType, String firstFormula, String secondFormula) {
		constraint = dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		return self();
	}

	public B createDecimalConstraint(int operatorType, String firstFormula, String secondFormula) {
		constraint = dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		return self();
	}

	public B createTextLengthConstraint(int operatorType, String firstFormula, String secondFormula) {
		constraint = dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		return self();
	}

	public B allowedEmptyCell(boolean allowedEmptyCell) {
		this.allowedEmptyCell = allowedEmptyCell;
		return self();
	}

	public B setErrorStyle(int errorStyle) {
		this.errorStyle = errorStyle;
		return self();
	}

	public B setSuppressDropDownArrow(boolean suppress) {
		this.suppress = suppress;
		return self();
	}

	public B showErrorBox(String errorBoxTitle, String errorBoxText) {
		showErrorBox(true, errorBoxTitle, errorBoxText);
		return self();
	}

	public B showPromptBox(String promptBoxTitle, String promptBoxText) {
		showPromptBox(true, promptBoxTitle, promptBoxText);
		return self();
	}

	protected B showErrorBox(boolean showErrorBox, String errorBoxTitle, String errorBoxText) {
		this.showErrorBox = showErrorBox;
		this.errorBoxTitle = errorBoxTitle;
		this.errorBoxText = errorBoxText;
		return self();
	}

	protected B showPromptBox(boolean showPromptBox, String promptBoxTitle, String promptBoxText) {
		this.showPromptBox = showPromptBox;
		this.promptBoxTitle = promptBoxTitle;
		this.promptBoxText = promptBoxText;
		return self();
	}

	@Override
	@SuppressWarnings("unchecked")
	protected B self() {
		return (B) this;
	}

	@Override
	public P end() {
		if (constraint != null) {
			CellRangeAddressList regions = new CellRangeAddressList();
			regions.addCellRangeAddress(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
			DataValidation validation = dataValidationHelper.createValidation(constraint, regions);
			validation.setEmptyCellAllowed(allowedEmptyCell);
			validation.setSuppressDropDownArrow(suppress);
			validation.setErrorStyle(errorStyle);
			validation.setShowErrorBox(showErrorBox);
			validation.setShowPromptBox(showPromptBox);
			validation.createErrorBox(errorBoxTitle, errorBoxText);
			validation.createPromptBox(promptBoxTitle, promptBoxText);
			parent.sheet.addValidationData(validation);
		}
		return super.end();
	}
}
