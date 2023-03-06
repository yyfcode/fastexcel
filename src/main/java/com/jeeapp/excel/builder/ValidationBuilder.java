package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;

/**
 * @author Justice
 */
public class ValidationBuilder<B extends CreationBuilder<B, P>, P extends SheetBuilderHelper<P>> {

	private final CreationBuilder<B, P> parent;

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

	protected ValidationBuilder(CreationBuilder<B, P> parent, DataValidation validation) {
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

	public CreationBuilder<B, P> addValidationData() {
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
