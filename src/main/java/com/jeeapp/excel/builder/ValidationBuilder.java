package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;

/**
 * @author Justice
 */
public class ValidationBuilder<B extends CreationBuilder<B>> {

	private final B parent;

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

	protected ValidationBuilder(B parent, DataValidation validation) {
		this.parent = parent;
		this.validation = validation;
	}

	public ValidationBuilder<B> allowedEmptyCell(boolean allowedEmptyCell) {
		this.allowedEmptyCell = allowedEmptyCell;
		return this;
	}

	public ValidationBuilder<B> setErrorStyle(int errorStyle) {
		this.errorStyle = errorStyle;
		return this;
	}

	public ValidationBuilder<B> setSuppressDropDownArrow(boolean suppress) {
		this.suppress = suppress;
		return this;
	}

	public ValidationBuilder<B> showErrorBox(String errorBoxTitle, String errorBoxText) {
		this.showErrorBox = true;
		this.errorBoxTitle = errorBoxTitle;
		this.errorBoxText = errorBoxText;
		return this;
	}

	public ValidationBuilder<B> showPromptBox(String promptBoxTitle, String promptBoxText) {
		this.showPromptBox = true;
		this.promptBoxTitle = promptBoxTitle;
		this.promptBoxText = promptBoxText;
		return this;
	}

	@Deprecated
	protected ValidationBuilder<B> showErrorBox(boolean showErrorBox, String errorBoxTitle, String errorBoxText) {
		this.showErrorBox = showErrorBox;
		this.errorBoxTitle = errorBoxTitle;
		this.errorBoxText = errorBoxText;
		return this;
	}

	@Deprecated
	protected ValidationBuilder<B> showPromptBox(boolean showPromptBox, String promptBoxTitle, String promptBoxText) {
		this.showPromptBox = showPromptBox;
		this.promptBoxTitle = promptBoxTitle;
		this.promptBoxText = promptBoxText;
		return this;
	}

	public B addValidationData() {
		validation.setEmptyCellAllowed(allowedEmptyCell);
		validation.setSuppressDropDownArrow(suppress);
		validation.setErrorStyle(errorStyle);
		validation.setShowErrorBox(showErrorBox);
		validation.setShowPromptBox(showPromptBox);
		validation.createErrorBox(errorBoxTitle, errorBoxText);
		validation.createPromptBox(promptBoxTitle, promptBoxText);
		parent.parent.sheet.addValidationData(validation);
		return parent.self();
	}
}
