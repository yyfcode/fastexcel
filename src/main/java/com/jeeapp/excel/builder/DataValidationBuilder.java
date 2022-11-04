package com.jeeapp.excel.builder;

import java.util.Arrays;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.util.Assert;

/**
 * @author Justice
 */
public class DataValidationBuilder {

	private final SheetBuilder parent;

	private final Sheet sheet;

	private final CellRangeAddressList region;

	private final DataValidationHelper dataValidationHelper;

	private DataValidationConstraint constraint;

	private int validationType;

	private int operatorType;

	private String firstFormula;

	private String secondFormula;

	private String[] explicitListValues;

	private boolean allowedEmptyCell = true;

	private boolean showPromptBox = false;

	private boolean showErrorBox = false;

	private int errorStyle = ErrorStyle.WARNING;

	private String errorBoxTitle;

	private String errorBoxText;

	private String promptBoxTitle;

	private String promptBoxText;

	public DataValidationBuilder(SheetBuilder parent, Sheet sheet, CellRangeAddressList region) {
		this.parent = parent;
		this.sheet = sheet;
		this.region = region;
		this.dataValidationHelper = sheet.getDataValidationHelper();
	}

	protected DataValidationBuilder createConstraint(int validationType, int operatorType, String firstFormula,
		String secondFormula, String[] explicitListValues, String dateFormat) {
		this.validationType = validationType;
		this.operatorType = operatorType;
		this.firstFormula = firstFormula;
		this.secondFormula = secondFormula;
		this.explicitListValues = explicitListValues;
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
		return this;
	}

	public DataValidationBuilder createExplicitListConstraint(String[] explicitListValues) {
		this.validationType = ValidationType.LIST;
		this.explicitListValues = explicitListValues;
		constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
		return this;
	}

	public DataValidationBuilder createFormulaListConstraint(String[] explicitListValues) {
		this.validationType = ValidationType.LIST;
		this.explicitListValues = explicitListValues;
		constraint = dataValidationHelper.createExplicitListConstraint(explicitListValues);
		return this;
	}

	public DataValidationBuilder createTimeConstraint(String firstFormula) {
		this.validationType = ValidationType.TIME;
		this.firstFormula = firstFormula;
		constraint = dataValidationHelper.createFormulaListConstraint(firstFormula);
		return this;
	}

	public DataValidationBuilder createDateConstraint(int operatorType, String firstFormula, String secondFormula, String dateFormat) {
		this.validationType = ValidationType.DATE;
		this.operatorType = operatorType;
		this.firstFormula = firstFormula;
		this.secondFormula = secondFormula;
		constraint = dataValidationHelper.createDateConstraint(operatorType, firstFormula, secondFormula, dateFormat);
		return this;
	}

	public DataValidationBuilder createCustomConstraint(String firstFormula) {
		this.validationType = ValidationType.FORMULA;
		this.firstFormula = firstFormula;
		constraint = dataValidationHelper.createCustomConstraint(firstFormula);
		return this;
	}

	public DataValidationBuilder createIntegerConstraint(int operatorType, String firstFormula, String secondFormula) {
		this.validationType = ValidationType.INTEGER;
		this.operatorType = operatorType;
		this.firstFormula = firstFormula;
		this.secondFormula = secondFormula;
		constraint = dataValidationHelper.createIntegerConstraint(operatorType, firstFormula, secondFormula);
		return this;
	}

	public DataValidationBuilder createDecimalConstraint(int operatorType, String firstFormula, String secondFormula) {
		this.validationType = ValidationType.DECIMAL;
		this.operatorType = operatorType;
		this.firstFormula = firstFormula;
		this.secondFormula = secondFormula;
		constraint = dataValidationHelper.createDecimalConstraint(operatorType, firstFormula, secondFormula);
		return this;
	}

	public DataValidationBuilder createTextLengthConstraint(int operatorType, String firstFormula, String secondFormula) {
		this.validationType = ValidationType.TEXT_LENGTH;
		this.operatorType = operatorType;
		this.firstFormula = firstFormula;
		this.secondFormula = secondFormula;
		constraint = dataValidationHelper.createTextLengthConstraint(operatorType, firstFormula, secondFormula);
		return this;
	}

	public DataValidationBuilder allowedEmptyCell(boolean allowedEmptyCell) {
		this.allowedEmptyCell = allowedEmptyCell;
		return this;
	}

	public DataValidationBuilder setErrorStyle(int errorStyle) {
		this.errorStyle = errorStyle;
		return this;
	}

	public DataValidationBuilder showErrorBox(String errorBoxTitle, String errorBoxText) {
		this.showErrorBox = true;
		this.errorBoxTitle = errorBoxTitle;
		this.errorBoxText = errorBoxText;
		return this;
	}

	public DataValidationBuilder showPromptBox(String promptBoxTitle, String promptBoxText) {
		this.showPromptBox = true;
		this.promptBoxTitle = promptBoxTitle;
		this.promptBoxText = promptBoxText;
		return this;
	}

	protected DataValidationBuilder showErrorBox(boolean showErrorBox, String errorBoxTitle, String errorBoxText) {
		this.showErrorBox = showErrorBox;
		this.errorBoxTitle = errorBoxTitle;
		this.errorBoxText = errorBoxText;
		return this;
	}

	protected DataValidationBuilder showPromptBox(boolean showPromptBox, String promptBoxTitle, String promptBoxText) {
		this.showPromptBox = showPromptBox;
		this.promptBoxTitle = promptBoxTitle;
		this.promptBoxText = promptBoxText;
		return this;
	}

	public DataValidationBuilder createValidation(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.addValidationData();
		return parent.createValidation(firstRow, lastRow, firstCol, lastCol);
	}

	public SheetBuilder addValidationData() {
		Assert.notNull(constraint, "Constraint must not be null!");
		DataValidation validation = dataValidationHelper.createValidation(constraint, region);
		validation.setEmptyCellAllowed(allowedEmptyCell);
		validation.setErrorStyle(errorStyle);
		if (showErrorBox) {
			validation.setShowErrorBox(true);
			if (StringUtils.isBlank(errorBoxText)) {
				errorBoxText = createDefaultText(true);
			}
			validation.createErrorBox(errorBoxTitle, errorBoxText);
		}
		if (showPromptBox) {
			validation.setShowPromptBox(true);
			if (StringUtils.isBlank(promptBoxText)) {
				promptBoxText = createDefaultText(false);
			}
			validation.createPromptBox(promptBoxTitle, promptBoxText);
		}
		sheet.addValidationData(validation);
		return parent;
	}

	protected String createDefaultText(boolean error) {
		String type = "";
		if (validationType == ValidationType.LIST) {
			if (explicitListValues != null) {
				return String.format("必须是%s其中之一", Arrays.toString(explicitListValues));
			} else {
				return String.format("必须符合验证规则：%s!", firstFormula);
			}
		}
		if (validationType == ValidationType.TIME) {
			type = "时间";
		}
		if (validationType == ValidationType.DATE) {
			type = "日期";
		}
		if (validationType == ValidationType.FORMULA) {
			return String.format("必须符合验证规则：%s!", firstFormula);
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
		return error ? "数据验证不通过!" : "";
	}
}
