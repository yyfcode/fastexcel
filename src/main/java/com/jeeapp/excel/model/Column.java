package com.jeeapp.excel.model;

import java.lang.reflect.Field;
import java.util.List;
import java.util.regex.Pattern;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.core.annotation.AnnotationUtils;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.annotation.ExcelProperty.Header;
import com.jeeapp.excel.annotation.ExcelProperty.Validation;

/**
 * @author Justice
 */
@Getter
@Setter
@ToString
@EqualsAndHashCode(callSuper = true)
public class Column extends Cell {

	private static final long serialVersionUID = 3129761200969567778L;

	private static Pattern SPLIT_CAMEL_CASE = Pattern.compile("(?<!(^|[A-Z]))(?=[A-Z])|(?<!^)(?=[A-Z][a-z])");

	protected final static Integer DEFAULT_WIDTH = 10;

	protected final static String DEFAULT_FORMAT = "General";

	/*-------------------------------------------
    |               列宽和格式化                  |
    ============================================*/

	private Integer width = DEFAULT_WIDTH;

	private String format = DEFAULT_FORMAT;

	/*-------------------------------------------
    |               列数据验证                    |
    ============================================*/

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

	/*-------------------------------------------
    |                表头属性                    |
    ============================================*/

	private List<Column> children;

	private BorderStyle border = BorderStyle.THIN;

	private short borderColor;

	private FillPatternType fillPatternType = FillPatternType.SOLID_FOREGROUND;

	private short fillBackgroundColor = IndexedColors.WHITE1.getIndex();

	private short fillForegroundColor = IndexedColors.GREY_25_PERCENT.getIndex();

	private short fontColor = IndexedColors.BLACK.getIndex();

	private String comment;

	private String commentAuthor;

	private Integer commentWidth;

	private Integer commentHeight;

	protected Column(String name, Object value) {
		super(name, value);
	}

	public static Column of(String name, Field field) {
		Column column = new Column(name, StringUtils.capitalize(StringUtils.join(SPLIT_CAMEL_CASE.split(field.getName()), " ")));
		ExcelProperty property = AnnotationUtils.getAnnotation(field, ExcelProperty.class);
		if (property == null) {
			return column;
		}
		column.setValue(property.name());
		column.setWidth(property.width());
		column.setFormat(property.format());

		// 数据验证
		Validation validation = property.validation();
		column.setValidationType(validation.validationType());
		column.setOperatorType(validation.operatorType());
		column.setFirstFormula(validation.firstFormula());
		column.setSecondFormula(validation.secondFormula());
		column.setExplicitListValues(validation.explicitListValues());
		column.setDateFormat(validation.dateFormat());
		column.setAllowEmpty(validation.allowEmpty());
		column.setErrorStyle(validation.errorStyle());
		column.setShowPromptBox(validation.showPromptBox());
		column.setPromptBoxTitle(validation.promptBoxTitle());
		column.setPromptBoxText(validation.promptBoxText());
		column.setShowErrorBox(validation.showErrorBox());
		column.setErrorBoxTitle(validation.errorBoxTitle());
		column.setErrorBoxText(validation.errorBoxText());

		// 表头
		Header header = property.header();
		column.setBorder(header.border());
		column.setBorderColor(header.borderColor().getIndex());
		column.setFillPatternType(header.fillPatternType());
		column.setFillForegroundColor(header.fillForegroundColor().getIndex());
		column.setFillBackgroundColor(header.fillBackgroundColor().getIndex());
		column.setFontColor(header.fontColor().getIndex());
		column.setComment(header.comment().value());
		column.setCommentAuthor(header.comment().author());
		column.setCommentWidth(header.comment().width());
		column.setCommentHeight(header.comment().height());

		return column;
	}

	public boolean hasComment() {
		return StringUtils.isNotBlank(this.getComment());
	}

	public boolean hasValidation() {
		return validationType >= 0;
	}
}
