package com.jeeapp.excel.model;

import java.io.Serializable;
import java.lang.reflect.Field;

import lombok.Data;
import org.springframework.core.annotation.AnnotationUtils;
import org.springframework.util.Assert;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.annotation.ExcelProperty.Validation;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author Justice
 */
@Data
public class Cell implements Serializable {

	private static final long serialVersionUID = 8481559629638325751L;

	protected final ExcelProperty property;

	/** 单元格对应的字段属性 */
	private String name;

	/** 单元格对应的字段值 */
	private Object value;

	/** 单元格起始行 */
	private int firstRow;

	/** 单元格结束行 */
	private int lastRow;

	/** 单元格起始列 */
	private int firstCol;

	/** 单元格结束列 */
	private int lastCol;

	/** 单元格格式 */
	private String format = CellUtils.DEFAULT_FORMAT;

	private int validationType;

	private int operatorType;

	private String firstFormula;

	private String secondFormula;

	private String[] explicitListValues;

	private boolean allowEmpty;

	private String dateFormat;

	private int errorStyle;

	private boolean showPromptBox;

	private String promptBoxTitle;

	private String promptBoxText;

	private boolean showErrorBox;

	private String errorBoxTitle;

	private String errorBoxText;

	public Cell(Field field) {
		this.property = AnnotationUtils.getAnnotation(field, ExcelProperty.class);
		Assert.notNull(property, String.format("@ExcelProperty not found for %s", field));
		// 数据格式
		this.setFormat(property.format());
		// 数据验证
		Validation validation = property.validation();
		this.setValidationType(validation.validationType());
		this.setOperatorType(validation.operatorType());
		this.setFirstFormula(validation.firstFormula());
		this.setSecondFormula(validation.secondFormula());
		this.setExplicitListValues(validation.explicitListValues());
		this.setDateFormat(validation.dateFormat());
		this.setAllowEmpty(validation.allowEmpty());
		this.setErrorStyle(validation.errorStyle());
		this.setShowPromptBox(validation.showPromptBox());
		this.setPromptBoxTitle(validation.promptBoxTitle());
		this.setPromptBoxText(validation.promptBoxText());
		this.setShowErrorBox(validation.showErrorBox());
		this.setErrorBoxTitle(validation.errorBoxTitle());
		this.setErrorBoxText(validation.errorBoxText());
	}
}
