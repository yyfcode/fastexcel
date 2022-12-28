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
import com.jeeapp.excel.annotation.ExcelProperty.Header;
import com.jeeapp.excel.annotation.ExcelProperty.Comment;

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

	/*-------------------------------------------
    |               列宽                        |
    ============================================*/

	private Integer width = DEFAULT_WIDTH;

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

	public Column(Field field) {
		super(field);
		this.setValue(StringUtils.capitalize(StringUtils.join(SPLIT_CAMEL_CASE.split(field.getName()), " ")));
		ExcelProperty property = AnnotationUtils.getAnnotation(field, ExcelProperty.class);
		Assert.notNull(property, String.format("@ExcelProperty not found for %s", field));
		// 标题和宽度
		this.setValue(property.name());
		this.setWidth(property.width());
		// 表头样式
		Header header = property.header();
		this.setBorder(header.border());
		this.setBorderColor(header.borderColor().getIndex());
		this.setFillPatternType(header.fillPatternType());
		this.setFillForegroundColor(header.fillForegroundColor().getIndex());
		this.setFillBackgroundColor(header.fillBackgroundColor().getIndex());
		this.setFontColor(header.fontColor().getIndex());
		// 表头批注
		Comment comment = header.comment();
		this.setComment(comment.value());
		this.setCommentAuthor(comment.author());
		this.setCommentWidth(comment.width());
		this.setCommentHeight(comment.height());
	}
}
