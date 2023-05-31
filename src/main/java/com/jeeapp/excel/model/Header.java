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
import com.jeeapp.excel.annotation.ExcelProperty;

/**
 * @author Justice
 */
@Getter
@Setter
@ToString
@EqualsAndHashCode(callSuper = true)
public class Header extends Column {

	private static Pattern SPLIT_CAMEL_CASE = Pattern.compile("(?<!(^|[A-Z]))(?=[A-Z])|(?<!^)(?=[A-Z][a-z])");

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

	public Header(Field field) {
		super(field);
		// 表头标题
		if (StringUtils.isBlank(property.name())) {
			this.setValue(StringUtils.capitalize(StringUtils.join(SPLIT_CAMEL_CASE.split(field.getName()), " ")));
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
}
