package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Justice
 */
public class DataValidationBuilder<P> {

	private final P parent;

	private final Sheet sheet;

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

	protected DataValidationBuilder(P parent, Sheet sheet, DataValidation validation) {
		this.parent = parent;
		this.sheet = sheet;
		this.validation = validation;
	}

	public DataValidationBuilder<P> allowedEmptyCell(boolean allowedEmptyCell) {
		this.allowedEmptyCell = allowedEmptyCell;
		return this;
	}

	public DataValidationBuilder<P> setErrorStyle(int errorStyle) {
		this.errorStyle = errorStyle;
		return this;
	}

	public DataValidationBuilder<P> setSuppressDropDownArrow(boolean suppress) {
		this.suppress = suppress;
		return this;
	}

	public DataValidationBuilder<P> showErrorBox(String errorBoxTitle, String errorBoxText) {
		showErrorBox(true, errorBoxTitle, errorBoxText);
		return this;
	}

	public DataValidationBuilder<P> showPromptBox(String promptBoxTitle, String promptBoxText) {
		showPromptBox(true, promptBoxTitle, promptBoxText);
		return this;
	}

	protected DataValidationBuilder<P> showErrorBox(boolean showErrorBox, String errorBoxTitle, String errorBoxText) {
		this.showErrorBox = showErrorBox;
		this.errorBoxTitle = errorBoxTitle;
		this.errorBoxText = errorBoxText;
		return this;
	}

	protected DataValidationBuilder<P> showPromptBox(boolean showPromptBox, String promptBoxTitle, String promptBoxText) {
		this.showPromptBox = showPromptBox;
		this.promptBoxTitle = promptBoxTitle;
		this.promptBoxText = promptBoxText;
		return this;
	}


	public P addValidationData() {
		validation.setEmptyCellAllowed(allowedEmptyCell);
		validation.setSuppressDropDownArrow(suppress);
		validation.setErrorStyle(errorStyle);
		validation.setShowErrorBox(showErrorBox);
		validation.setShowPromptBox(showPromptBox);
		validation.createErrorBox(errorBoxTitle, errorBoxText);
		validation.createPromptBox(promptBoxTitle, promptBoxText);
		sheet.addValidationData(validation);
		return parent;
	}
}
