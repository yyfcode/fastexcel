package com.jeeapp.excel.builder;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Predicate;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author justice
 */
public class CellStyleBuilder<B extends CellStyleBuilder<B, P>, P extends CellBuilderHelper<P>> {

	private final P parent;

	private final Map<String, Object> properties;

	private Predicate<Cell> predicate;

	private Integer column;

	protected CellRangeAddress region;

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

	protected CellStyleBuilder(P parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		this.parent = parent;
		this.region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		this.predicate = cell -> cell.getColumnIndex() >= region.getFirstColumn()
			&& cell.getColumnIndex() <= region.getLastColumn()
			&& cell.getRowIndex() >= region.getFirstRow()
			&& cell.getRowIndex() <= region.getLastRow();
		this.properties = new HashMap<>();
	}

	/**
	 * 字符集
	 */
	public B setCharSet(byte charSet) {
		properties.put(CellUtils.CHAR_SET, charSet);
		return self();
	}

	/**
	 * 斜体
	 */
	public B setItalic(boolean italic) {
		properties.put(CellUtils.ITALIC, italic);
		return self();
	}

	/**
	 * 字体
	 */
	public B setFontName(String fontName) {
		properties.put(CellUtils.FONT_NAME, fontName);
		return self();
	}

	/**
	 * 偏移量
	 */
	public B setTypeOffset(byte typeOffset) {
		properties.put(CellUtils.TYPE_OFFSET, typeOffset);
		return self();
	}

	/**
	 * 下划线
	 */
	public B setUnderline(byte underline) {
		properties.put(CellUtils.UNDERLINE, underline);
		return self();
	}

	/**
	 * 字体大小
	 */
	public B setFontHeight(int fontHeight) {
		properties.put(CellUtils.FONT_HEIGHT, (short) (fontHeight * 20));
		return self();
	}

	/**
	 * 删除线
	 */
	public B setStrikeout(boolean strikeout) {
		properties.put(CellUtils.STRIKEOUT, strikeout);
		return self();
	}

	/**
	 * 粗体
	 */
	public B setFontBold(boolean bold) {
		properties.put(CellUtils.BOLD, bold);
		return self();
	}

	/**
	 * 字体颜色
	 */
	public B setFontColor(IndexedColors color) {
		properties.put(CellUtils.COLOR, color.getIndex());
		return self();
	}

	/**
	 * 字体颜色
	 */
	public B setFontColor(short color) {
		properties.put(CellUtils.COLOR, color);
		return self();
	}

	/**
	 * 隐藏
	 */
	public B setHidden(boolean hidden) {
		properties.put(CellUtil.HIDDEN, hidden);
		return self();
	}

	/**
	 * 锁定
	 */
	public B setLocked(boolean locked) {
		properties.put(CellUtil.LOCKED, locked);
		return self();
	}

	/**
	 * 缩进
	 */
	public B setIndention(short indent) {
		properties.put(CellUtil.INDENTION, indent);
		return self();
	}

	/**
	 * 旋转
	 */
	public B setRotation(short rotation) {
		properties.put(CellUtil.ROTATION, rotation);
		return self();
	}

	/**
	 * 换行
	 */
	public B setWrapText(boolean wrapText) {
		properties.put(CellUtil.WRAP_TEXT, wrapText);
		return self();
	}

	/**
	 * 横向位置
	 */
	public B setAlignment(HorizontalAlignment alignment) {
		properties.put(CellUtil.ALIGNMENT, alignment);
		return self();
	}

	/**
	 * 纵向位置
	 */
	public B setVerticalAlignment(VerticalAlignment alignment) {
		properties.put(CellUtil.VERTICAL_ALIGNMENT, alignment);
		return self();
	}

	/**
	 * 线框样式
	 */
	public B setBorder(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_TOP, borderStyle);
		properties.put(CellUtil.BORDER_BOTTOM, borderStyle);
		properties.put(CellUtil.BORDER_LEFT, borderStyle);
		properties.put(CellUtil.BORDER_RIGHT, borderStyle);
		return self();
	}

	/**
	 * 上边线框样式
	 */
	public B setBorderTop(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_TOP, borderStyle);
		return self();
	}

	/**
	 * 左边线框样式
	 */
	public B setBorderLeft(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_LEFT, borderStyle);
		return self();
	}

	/**
	 * 下边线框样式
	 */
	public B setBorderBottom(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_BOTTOM, borderStyle);
		return self();
	}

	/**
	 * 右边线框样式
	 */
	public B setBorderRight(BorderStyle borderStyle) {
		properties.put(CellUtil.BORDER_RIGHT, borderStyle);
		return self();
	}

	/**
	 * 线框颜色
	 */
	public B setBorderColor(IndexedColors color) {
		short index = color.getIndex();
		properties.put(CellUtil.TOP_BORDER_COLOR, index);
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, index);
		properties.put(CellUtil.LEFT_BORDER_COLOR, index);
		properties.put(CellUtil.RIGHT_BORDER_COLOR, index);
		return self();
	}

	/**
	 * 线框颜色
	 */
	public B setBorderColor(short color) {
		properties.put(CellUtil.TOP_BORDER_COLOR, color);
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, color);
		properties.put(CellUtil.LEFT_BORDER_COLOR, color);
		properties.put(CellUtil.RIGHT_BORDER_COLOR, color);
		return self();
	}

	/**
	 * 上边线框颜色
	 */
	public B setTopBorderColor(IndexedColors color) {
		properties.put(CellUtil.TOP_BORDER_COLOR, color.getIndex());
		return self();
	}

	/**
	 * 上边线框颜色
	 */
	public B setTopBorderColor(short color) {
		properties.put(CellUtil.TOP_BORDER_COLOR, color);
		return self();
	}

	/**
	 * 左边线框颜色
	 */
	public B setLeftBorderColor(IndexedColors color) {
		properties.put(CellUtil.LEFT_BORDER_COLOR, color.getIndex());
		return self();
	}

	/**
	 * 左边线框颜色
	 */
	public B setLeftBorderColor(short color) {
		properties.put(CellUtil.LEFT_BORDER_COLOR, color);
		return self();
	}

	/**
	 * 下边线框颜色
	 */
	public B setBottomBorderColor(IndexedColors color) {
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, color.getIndex());
		return self();
	}

	/**
	 * 下边线框颜色
	 */
	public B setBottomBorderColor(short color) {
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, color);
		return self();
	}

	/**
	 * 右边线框颜色
	 */
	public B setRightBorderColor(IndexedColors color) {
		properties.put(CellUtil.RIGHT_BORDER_COLOR, color.getIndex());
		return self();
	}

	/**
	 * 右边线框颜色
	 */
	public B setRightBorderColor(short color) {
		properties.put(CellUtil.RIGHT_BORDER_COLOR, color);
		return self();
	}

	/**
	 * 背景色
	 */
	public B setFillBackgroundColor(IndexedColors color) {
		properties.put(CellUtil.FILL_BACKGROUND_COLOR, color.getIndex());
		return self();
	}

	/**
	 * 背景色
	 */
	public B setFillBackgroundColor(short color) {
		properties.put(CellUtil.FILL_BACKGROUND_COLOR, color);
		return self();
	}

	/**
	 * 前景色
	 */
	public B setFillForegroundColor(IndexedColors color) {
		properties.put(CellUtil.FILL_FOREGROUND_COLOR, color.getIndex());
		return self();
	}

	/**
	 * 前景色
	 */
	public B setFillForegroundColor(short color) {
		properties.put(CellUtil.FILL_FOREGROUND_COLOR, color);
		return self();
	}

	/**
	 * 填充类型
	 */
	public B setFillPattern(FillPatternType fp) {
		properties.put(CellUtil.FILL_PATTERN, fp);
		return self();
	}

	/**
	 * 格式化
	 */
	public B setDataFormat(String pFmt) {
		properties.put(CellUtil.DATA_FORMAT, parent.workbook.createDataFormat().getFormat(pFmt));
		return self();
	}

	@SuppressWarnings("unchecked")
	protected B self() {
		return (B) this;
	}

	public Workbook build() {
		return end().build();
	}

	public P end() {
		if (properties.isEmpty()) {
			return parent;
		}
		if (region != null) {
			parent.addCellStyle(predicate, properties);
			parent.addRegionStyle(region, properties);
		} else if (predicate != null) {
			parent.addCellStyle(predicate, properties);
		} else if (column != null) {
			parent.addColumnStyle(column, properties);
		} else {
			parent.addCommonStyle(properties);
		}
		return parent;
	}

	/**
	 * @deprecated use {@link CellStyleBuilder#end()} instead.
	 */
	@Deprecated
	public P addCellStyle() {
		return end();
	}
}
