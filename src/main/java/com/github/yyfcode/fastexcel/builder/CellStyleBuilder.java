package com.github.yyfcode.fastexcel.builder;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Predicate;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellUtil;
import com.github.yyfcode.fastexcel.util.CellUtils;

/**
 * @author justice
 */
public class CellStyleBuilder<P extends CellBuilderHelper<P>> {

	private final P parent;

	private Predicate<Cell> predicate;

	private Integer column;

	private final Map<String, Object> properties;

	protected CellStyleBuilder(P parent) {
		this.parent = parent;
		this.properties = new HashMap<>();
	}

	protected CellStyleBuilder(P parent, int column) {
		this.parent = parent;
		this.column = column;
		this.properties = new HashMap<>();
	}


	protected CellStyleBuilder(P parent, Predicate<Cell> predicate) {
		this.parent = parent;
		this.predicate = predicate;
		this.properties = new HashMap<>();
	}

	/**
	 * 字符集
	 */
	public CellStyleBuilder<P> setCharSet(byte charSet) {
		properties.put(CellUtils.CHAR_SET, charSet);
		return this;
	}

	/**
	 * 斜体
	 */
	public CellStyleBuilder<P> setItalic(boolean italic) {
		properties.put(CellUtils.ITALIC, italic);
		return this;
	}

	/**
	 * 字体
	 */
	public CellStyleBuilder<P> setFontName(String fontName) {
		properties.put(CellUtils.FONT_NAME, fontName);
		return this;
	}

	/**
	 * 偏移量
	 */
	public CellStyleBuilder<P> setTypeOffset(byte typeOffset) {
		properties.put(CellUtils.TYPE_OFFSET, typeOffset);
		return this;
	}

	/**
	 * 下划线
	 */
	public CellStyleBuilder<P> setUnderline(byte underline) {
		properties.put(CellUtils.UNDERLINE, underline);
		return this;
	}

	/**
	 * 字体大小
	 */
	public CellStyleBuilder<P> setFontHeight(int fontHeight) {
		properties.put(CellUtils.FONT_HEIGHT, (short) (fontHeight * 20));
		return this;
	}

	/**
	 * 删除线
	 */
	public CellStyleBuilder<P> setStrikeout(boolean strikeout) {
		properties.put(CellUtils.STRIKEOUT, strikeout);
		return this;
	}

	/**
	 * 粗体
	 */
	public CellStyleBuilder<P> setFontBold(boolean bold) {
		properties.put(CellUtils.BOLD, bold);
		return this;
	}

	/**
	 * 字体颜色
	 */
	public CellStyleBuilder<P> setFontColor(IndexedColors color) {
		properties.put(CellUtils.COLOR, color.getIndex());
		return this;
	}

	/**
	 * 字体颜色
	 */
	public CellStyleBuilder<P> setFontColor(short color) {
		properties.put(CellUtils.COLOR, color);
		return this;
	}

	/**
	 * 隐藏
	 */
	public CellStyleBuilder<P> setHidden(boolean hidden) {
		properties.put(CellUtil.HIDDEN, hidden);
		return this;
	}

	/**
	 * 锁定
	 */
	public CellStyleBuilder<P> setLocked(boolean locked) {
		properties.put(CellUtil.LOCKED, locked);
		return this;
	}

	/**
	 * 缩进
	 */
	public CellStyleBuilder<P> setIndention(short indent) {
		properties.put(CellUtil.INDENTION, indent);
		return this;
	}

	/**
	 * 旋转
	 */
	public CellStyleBuilder<P> setRotation(short rotation) {
		properties.put(CellUtil.ROTATION, rotation);
		return this;
	}

	/**
	 * 换行
	 */
	public CellStyleBuilder<P> setWrapText(boolean wrapText) {
		properties.put(CellUtil.WRAP_TEXT, wrapText);
		return this;
	}

	/**
	 * 横向位置
	 */
	public CellStyleBuilder<P> setAlignment(HorizontalAlignment alignment) {
		properties.put(CellUtil.ALIGNMENT, alignment);
		return this;
	}

	/**
	 * 纵向位置
	 */
	public CellStyleBuilder<P> setVerticalAlignment(VerticalAlignment alignment) {
		properties.put(CellUtil.VERTICAL_ALIGNMENT, alignment);
		return this;
	}

	/**
	 * 线框样式
	 */
	public CellStyleBuilder<P> setBorder(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_TOP, borderStyle);
		properties.put(CellUtil.BORDER_BOTTOM, borderStyle);
		properties.put(CellUtil.BORDER_LEFT, borderStyle);
		properties.put(CellUtil.BORDER_RIGHT, borderStyle);
		return this;
	}

	/**
	 * 上边线框样式
	 */
	public CellStyleBuilder<P> setBorderTop(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_TOP, borderStyle);
		return this;
	}

	/**
	 * 左边线框样式
	 */
	public CellStyleBuilder<P> setBorderLeft(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_LEFT, borderStyle);
		return this;
	}

	/**
	 * 下边线框样式
	 */
	public CellStyleBuilder<P> setBorderBottom(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_BOTTOM, borderStyle);
		return this;
	}

	/**
	 * 右边线框样式
	 */
	public CellStyleBuilder<P> setBorderRight(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_RIGHT, borderStyle);
		return this;
	}

	/**
	 * 线框颜色
	 */
	public CellStyleBuilder<P> setBorderColor(IndexedColors color) {
		short index = color.getIndex();
		properties.put(CellUtil.TOP_BORDER_COLOR, index);
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, index);
		properties.put(CellUtil.LEFT_BORDER_COLOR, index);
		properties.put(CellUtil.RIGHT_BORDER_COLOR, index);
		return this;
	}

	/**
	 * 线框颜色
	 */
	public CellStyleBuilder<P> setBorderColor(short color) {
		properties.put(CellUtil.TOP_BORDER_COLOR, color);
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, color);
		properties.put(CellUtil.LEFT_BORDER_COLOR, color);
		properties.put(CellUtil.RIGHT_BORDER_COLOR, color);
		return this;
	}

	/**
	 * 上边线框颜色
	 */
	public CellStyleBuilder<P> setTopBorderColor(IndexedColors color) {
		properties.put(CellUtil.TOP_BORDER_COLOR, color.getIndex());
		return this;
	}

	/**
	 * 上边线框颜色
	 */
	public CellStyleBuilder<P> setTopBorderColor(short color) {
		properties.put(CellUtil.TOP_BORDER_COLOR, color);
		return this;
	}

	/**
	 * 左边线框颜色
	 */
	public CellStyleBuilder<P> setLeftBorderColor(IndexedColors color) {
		properties.put(CellUtil.LEFT_BORDER_COLOR, color.getIndex());
		return this;
	}

	/**
	 * 左边线框颜色
	 */
	public CellStyleBuilder<P> setLeftBorderColor(short color) {
		properties.put(CellUtil.LEFT_BORDER_COLOR, color);
		return this;
	}

	/**
	 * 下边线框颜色
	 */
	public CellStyleBuilder<P> setBottomBorderColor(IndexedColors color) {
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, color.getIndex());
		return this;
	}

	/**
	 * 下边线框颜色
	 */
	public CellStyleBuilder<P> setBottomBorderColor(short color) {
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, color);
		return this;
	}

	/**
	 * 右边线框颜色
	 */
	public CellStyleBuilder<P> setRightBorderColor(IndexedColors color) {
		properties.put(CellUtil.RIGHT_BORDER_COLOR, color.getIndex());
		return this;
	}

	/**
	 * 右边线框颜色
	 */
	public CellStyleBuilder<P> setRightBorderColor(short color) {
		properties.put(CellUtil.RIGHT_BORDER_COLOR, color);
		return this;
	}

	/**
	 * 背景色
	 */
	public CellStyleBuilder<P> setFillBackgroundColor(IndexedColors color) {
		properties.put(CellUtil.FILL_BACKGROUND_COLOR, color.getIndex());
		return this;
	}

	/**
	 * 背景色
	 */
	public CellStyleBuilder<P> setFillBackgroundColor(short color) {
		properties.put(CellUtil.FILL_BACKGROUND_COLOR, color);
		return this;
	}

	/**
	 * 前景色
	 */
	public CellStyleBuilder<P> setFillForegroundColor(IndexedColors color) {
		properties.put(CellUtil.FILL_BACKGROUND_COLOR, color.getIndex());
		return this;
	}

	/**
	 * 前景色
	 */
	public CellStyleBuilder<P> setFillForegroundColor(short color) {
		properties.put(CellUtil.FILL_FOREGROUND_COLOR, color);
		return this;
	}

	/**
	 * 填充类型
	 */
	public CellStyleBuilder<P> setFillPattern(FillPatternType fp) {
		properties.put(CellUtil.FILL_PATTERN, fp);
		return this;
	}

	/**
	 * 格式化
	 */
	public CellStyleBuilder<P> setDataFormat(String pFmt) {
		properties.put(CellUtil.DATA_FORMAT, parent.createDataFormat().getFormat(pFmt));
		return this;
	}

	/**
	 * 添加样式
	 */
	public P addCellStyle() {
		if (predicate != null) {
			parent.addCellStyle(predicate, properties);
		} else if (column != null) {
			parent.addColumnStyle(column, properties);
		} else {
			parent.addCommonStyle(properties);
		}
		return parent;
	}

}
