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
import org.springframework.util.Assert;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.annotation.ExcelProperty.Validation;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author Justice
 */
@Getter
@Setter
@ToString
@EqualsAndHashCode(callSuper = true)
public class Header extends Cell {

	private static Pattern SPLIT_CAMEL_CASE = Pattern.compile("(?<!(^|[A-Z]))(?=[A-Z])|(?<!^)(?=[A-Z][a-z])");

	private final static Integer DEFAULT_WIDTH = 8;

	private Integer width = DEFAULT_WIDTH;

	private String format = CellUtils.DEFAULT_FORMAT;

	private Boolean hidden;

	private BorderStyle border = BorderStyle.THIN;

	private Short borderColor;

	private FillPatternType fillPatternType = FillPatternType.SOLID_FOREGROUND;

	private Short fillBackgroundColor = IndexedColors.WHITE1.getIndex();

	private Short fillForegroundColor = IndexedColors.GREY_25_PERCENT.getIndex();

	private Short fontColor = IndexedColors.BLACK.getIndex();

	private String comment;

	private String commentAuthor;

	private Integer commentWidth;

	private Integer commentHeight;

	private List<Header> children;

	@Deprecated
	private int validationType;

	@Deprecated
	private int operatorType;

	@Deprecated
	private String firstFormula;

	@Deprecated
	private String secondFormula;

	@Deprecated
	private String[] explicitListValues;

	@Deprecated
	private boolean allowEmpty;

	@Deprecated
	private String dateFormat;

	@Deprecated
	private int errorStyle;

	@Deprecated
	private boolean showPromptBox;

	@Deprecated
	private String promptBoxTitle;

	@Deprecated
	private String promptBoxText;

	@Deprecated
	private boolean showErrorBox;

	@Deprecated
	private String errorBoxTitle;

	@Deprecated
	private String errorBoxText;

	public Header(Field field) {
		ExcelProperty property = AnnotationUtils.getAnnotation(field, ExcelProperty.class);
		Assert.notNull(property, String.format("@ExcelProperty not found for %s", field));
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

		this.setFormat(property.format());
		this.setWidth(property.width());
		this.setHidden(property.hidden());
		// 表头标题
		if (StringUtils.isBlank(property.name())) {
			this.setValue(splitCamelCaseToString(field.getName()));
		} else {
			this.setValue(property.name());
		}

		// 表头样式
		ExcelProperty.Header header = property.header();
		this.setBorder(header.border());
		this.setBorderColor(header.borderColor().getIndex());
		this.setFillPatternType(header.fillPatternType());
		this.setFillForegroundColor(header.fillForegroundColor().getIndex());
		this.setFillBackgroundColor(header.fillBackgroundColor().getIndex());
		this.setFontColor(header.fontColor().getIndex());
		// 表头批注
		ExcelProperty.Comment comment = header.comment();
		this.setComment(comment.value());
		this.setCommentAuthor(comment.author());
		this.setCommentWidth(comment.width());
		this.setCommentHeight(comment.height());
	}

	private static String splitCamelCaseToString(String name) {
		return StringUtils.capitalize(StringUtils.join(SPLIT_CAMEL_CASE.split(name), " "));
	}
}
