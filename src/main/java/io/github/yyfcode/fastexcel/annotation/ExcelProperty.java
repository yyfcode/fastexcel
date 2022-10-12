package io.github.yyfcode.fastexcel.annotation;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author Justice
 */
@Inherited
@Target(FIELD)
@Retention(RUNTIME)
public @interface ExcelProperty {

	/**
	 * 列名
	 */
	String name();

	/**
	 * 列序
	 */
	int column();

	/**
	 * 列格式
	 */
	String format() default "General";

	/**
	 * 列宽
	 */
	int width() default 10;

	/**
	 * 表头样式
	 */
	Header header() default @Header;

	/**
	 * 数据验证
	 * @see org.apache.poi.ss.usermodel.DataValidation
	 */
	Validation validation() default @Validation;

	/**
	 * @author yinyf
	 */
	@Target({})
	@Retention(RUNTIME)
	@interface Header {

		/**
		 * 设置要用于单元格边框的边框类型
		 * @see BorderStyle#THIN
		 */
		BorderStyle border() default BorderStyle.THIN;

		/**
		 * 设置要用于边框的颜色
		 * @see IndexedColors#BLACK1
		 */
		IndexedColors borderColor() default IndexedColors.BLACK1;

		/**
		 * 指示用于单元格格式的填充样式的样式
		 * @see FillPatternType#SOLID_FOREGROUND
		 */
		FillPatternType fillPatternType() default FillPatternType.SOLID_FOREGROUND;

		/**
		 * 设置背景填充颜色
		 * @see IndexedColors#WHITE1
		 */
		IndexedColors fillBackgroundColor() default IndexedColors.WHITE1;

		/**
		 * 设置前景填充颜色 <i>注意：确保前景色设置在背景色之前</i>
		 * @see IndexedColors#GREY_25_PERCENT
		 */
		IndexedColors fillForegroundColor() default IndexedColors.GREY_25_PERCENT;

		/**
		 * 设置字体颜色
		 * @see IndexedColors#BLACK
		 */
		IndexedColors fontColor() default IndexedColors.BLACK;

		/**
		 * 设置表头批注
		 * @see org.apache.poi.ss.usermodel.Comment
		 */
		Comment comment() default @Comment;
	}

	/**
	 * @author justice
	 */
	@Target({})
	@Retention(RUNTIME)
	@interface Comment {

		/**
		 * 内容
		 */
		String value() default "";

		/**
		 * 作者
		 */
		String author() default "";

		/**
		 * 批注宽
		 */
		int width() default 2;

		/**
		 * 批注高
		 */
		int height() default 1;
	}

	/**
	 * @author Justice
	 */
	@Target({})
	@Retention(RUNTIME)
	@interface Validation {

		/**
		 * 验证类型
		 * @see org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType
		 */
		int validationType() default -1;

		/**
		 * 操作类型
		 * @see OperatorType
		 */
		int operatorType() default OperatorType.IGNORED;

		/**
		 * 表达式1
		 */
		String firstFormula() default "";

		/**
		 * 表达式2
		 */
		String secondFormula() default "";

		/**
		 * 显式列表值
		 */
		String[] explicitListValues() default {};

		/**
		 * 错误样式
		 * @see ErrorStyle
		 */
		int errorStyle() default ErrorStyle.WARNING;

		/**
		 * 显示错误提示
		 */
		boolean showErrorBox() default true;

		/**
		 * 错误提示标题
		 */
		String errorBoxTitle() default "";

		/**
		 * 错误提示内容
		 */
		String errorBoxText() default "";

		/**
		 * 显示填写提示
		 */
		boolean showPromptBox() default false;

		/**
		 * 填写提示标题
		 */
		String promptBoxTitle() default "";

		/**
		 * 填写提示内容
		 */
		String promptBoxText() default "";

		/**
		 * 是否必填
		 */
		boolean allowEmpty() default true;
	}
}
