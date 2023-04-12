package com.jeeapp.excel.util;

import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * @author justice
 */
@Slf4j
public class CellUtils {

	public static final String DEFAULT_FORMAT = "General";
	public static final String DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";
	public static final String CHAR_SET = "charSet";
	public static final String ITALIC = "italic";
	public static final String FONT_NAME = "fontName";
	public static final String TYPE_OFFSET = "typeOffset";
	public static final String UNDERLINE = "underline";
	public static final String FONT_HEIGHT = "fontHeight";
	public static final String STRIKEOUT = "strikeout";
	public static final String BOLD = "bold";
	public static final String COLOR = "color";

	private static final Set<String> fontValues = Collections.unmodifiableSet(
		new HashSet<>(Arrays.asList(
			FONT_NAME,
			UNDERLINE,
			TYPE_OFFSET,
			CHAR_SET,
			FONT_HEIGHT,
			COLOR,
			ITALIC,
			STRIKEOUT,
			BOLD
		)));

	private static final Set<String> shortValues = Collections.unmodifiableSet(
		new HashSet<>(Arrays.asList(
			CellUtil.BOTTOM_BORDER_COLOR,
			CellUtil.LEFT_BORDER_COLOR,
			CellUtil.RIGHT_BORDER_COLOR,
			CellUtil.TOP_BORDER_COLOR,
			CellUtil.FILL_FOREGROUND_COLOR,
			CellUtil.FILL_BACKGROUND_COLOR,
			CellUtil.INDENTION,
			CellUtil.DATA_FORMAT,
			CellUtil.ROTATION
		)));

	private static final Set<String> intValues = Collections.unmodifiableSet(
		new HashSet<>(Collections.singletonList(
			CellUtil.FONT
		)));

	private static final Set<String> booleanValues = Collections.unmodifiableSet(
		new HashSet<>(Arrays.asList(
			CellUtil.LOCKED,
			CellUtil.HIDDEN,
			CellUtil.WRAP_TEXT
		)));

	private static final Set<String> borderTypeValues = Collections.unmodifiableSet(
		new HashSet<>(Arrays.asList(
			CellUtil.BORDER_BOTTOM,
			CellUtil.BORDER_LEFT,
			CellUtil.BORDER_RIGHT,
			CellUtil.BORDER_TOP
		)));

	private static CellStyle getStyle(CellStyle originalStyle, Workbook workbook, Map<String, Object> properties) {
		CellStyle newStyle = null;
		Map<String, Object> values = getFormatProperties(originalStyle);
		properties.put(CellUtil.FONT, getFont(workbook, properties));
		putStyleProperties(properties, values);

		// index seems like what index the cellstyle is in the list of styles for a workbook.
		// not good to compare on!
		int numberCellStyles = workbook.getNumCellStyles();

		for (int i = 0; i < numberCellStyles; i++) {
			CellStyle wbStyle = workbook.getCellStyleAt(i);
			Map<String, Object> wbStyleMap = getFormatProperties(wbStyle);

			// the desired style already exists in the workbook. Use the existing style.
			if (wbStyleMap.equals(values)) {
				newStyle = wbStyle;
				break;
			}
		}

		// the desired style does not exist in the workbook. Create a new style with desired properties.
		if (newStyle == null) {
			newStyle = workbook.createCellStyle();
			setFormatProperties(newStyle, workbook, values);
		}
		return newStyle;
	}

	public static void setRegionStyleProperties(Sheet sheet, CellRangeAddress region, Map<String, Object> properties) {
		if (properties.containsKey(CellUtil.BORDER_LEFT)) {
			BorderStyle borderStyle = CellUtils.getBorderStyle(properties, CellUtil.BORDER_LEFT);
			RegionUtil.setBorderLeft(borderStyle, region, sheet);
		}
		if (properties.containsKey(CellUtil.BORDER_BOTTOM)) {
			BorderStyle borderStyle = CellUtils.getBorderStyle(properties, CellUtil.BORDER_BOTTOM);
			RegionUtil.setBorderBottom(borderStyle, region, sheet);
		}
		if (properties.containsKey(CellUtil.BORDER_RIGHT)) {
			BorderStyle borderStyle = CellUtils.getBorderStyle(properties, CellUtil.BORDER_RIGHT);
			RegionUtil.setBorderRight(borderStyle, region, sheet);
		}
		if (properties.containsKey(CellUtil.BORDER_TOP)) {
			BorderStyle borderStyle = CellUtils.getBorderStyle(properties, CellUtil.BORDER_TOP);
			RegionUtil.setBorderTop(borderStyle, region, sheet);
		}
		if (properties.containsKey(CellUtil.LEFT_BORDER_COLOR)) {
			short borderColor = CellUtils.getShort(properties, CellUtil.LEFT_BORDER_COLOR);
			RegionUtil.setLeftBorderColor(borderColor, region, sheet);
		}
		if (properties.containsKey(CellUtil.BOTTOM_BORDER_COLOR)) {
			short borderColor = CellUtils.getShort(properties, CellUtil.BOTTOM_BORDER_COLOR);
			RegionUtil.setRightBorderColor(borderColor, region, sheet);
		}
		if (properties.containsKey(CellUtil.RIGHT_BORDER_COLOR)) {
			short borderColor = CellUtils.getShort(properties, CellUtil.RIGHT_BORDER_COLOR);
			RegionUtil.setBottomBorderColor(borderColor, region, sheet);
		}
		if (properties.containsKey(CellUtil.TOP_BORDER_COLOR)) {
			short borderColor = CellUtils.getShort(properties, CellUtil.TOP_BORDER_COLOR);
			RegionUtil.setTopBorderColor(borderColor, region, sheet);
		}
	}

	public static void setRowStyleProperties(Sheet sheet, Row row, Map<String, Object> properties) {
		Workbook workbook = sheet.getWorkbook();
		CellStyle rowStyle = row.getRowStyle();
		if (!properties.isEmpty() && rowStyle == null) {
			rowStyle = workbook.createCellStyle();
		}
		row.setRowStyle(getStyle(rowStyle, workbook, properties));
	}

	public static void setCellStyleProperties(Cell cell, Map<String, Object> properties) {
		Workbook workbook = cell.getSheet().getWorkbook();
		CellStyle cellStyle = cell.getCellStyle();
		cell.setCellStyle(getStyle(cellStyle, workbook, properties));
	}

	public static void setColumnStyleProperties(Sheet sheet, int column, Map<String, Object> properties) {
		Workbook workbook = sheet.getWorkbook();
		CellStyle columnStyle = sheet.getColumnStyle(column);
		sheet.setDefaultColumnStyle(column, getStyle(columnStyle, workbook, properties));
	}

	private static Integer getFont(Workbook workbook, Map<String, Object> properties) {
		Font originalFont, newFont = null;
		if (workbook.getNumberOfFonts() > 0) {
			originalFont = workbook.getFontAt(0);
		} else {
			originalFont = workbook.createFont();
		}
		Map<String, Object> values = getFontProperties(originalFont);
		putFontProperties(properties, values);

		int numberFonts = workbook.getNumberOfFonts();
		for (int i = 0; i < numberFonts; i++) {
			Font wbFont = workbook.getFontAt(i);
			Map<String, Object> wbFontMap = getFontProperties(wbFont);
			if (wbFontMap.equals(values)) {
				newFont = wbFont;
				break;
			}
		}

		if (newFont == null) {
			newFont = workbook.createFont();
			setFontProperties(newFont, values);
		}
		return newFont.getIndex();
	}

	private static Map<String, Object> getFormatProperties(CellStyle style) {
		Map<String, Object> properties = new HashMap<>();
		put(properties, CellUtil.ALIGNMENT, style.getAlignment());
		put(properties, CellUtil.VERTICAL_ALIGNMENT, style.getVerticalAlignment());
		put(properties, CellUtil.BORDER_BOTTOM, style.getBorderBottom());
		put(properties, CellUtil.BORDER_LEFT, style.getBorderLeft());
		put(properties, CellUtil.BORDER_RIGHT, style.getBorderRight());
		put(properties, CellUtil.BORDER_TOP, style.getBorderTop());
		put(properties, CellUtil.BOTTOM_BORDER_COLOR, style.getBottomBorderColor());
		put(properties, CellUtil.DATA_FORMAT, style.getDataFormat());
		put(properties, CellUtil.FILL_PATTERN, style.getFillPattern());
		// The apache poi CellUtil only works using org.apache.poi.ss.*.
		// It cannot work using a XSSFColor because CellStyle has no method to get/set fill foreground color from a XSSFColor.
		// It only works using short color indexes from IndexedColors.
		// So if it's black here, change it to white
		short fillForegroundColor = style.getFillForegroundColor();
		put(properties, CellUtil.FILL_FOREGROUND_COLOR,
			fillForegroundColor == IndexedColors.AUTOMATIC.index ? IndexedColors.WHITE.index : fillForegroundColor);
		short fillBackgroundColor = style.getFillBackgroundColor();
		put(properties, CellUtil.FILL_BACKGROUND_COLOR,
			fillBackgroundColor == IndexedColors.AUTOMATIC.index ? IndexedColors.WHITE.index : fillBackgroundColor);
		put(properties, CellUtil.FONT, style.getFontIndex());
		put(properties, CellUtil.HIDDEN, style.getHidden());
		put(properties, CellUtil.INDENTION, style.getIndention());
		put(properties, CellUtil.LEFT_BORDER_COLOR, style.getLeftBorderColor());
		put(properties, CellUtil.LOCKED, style.getLocked());
		put(properties, CellUtil.RIGHT_BORDER_COLOR, style.getRightBorderColor());
		put(properties, CellUtil.ROTATION, style.getRotation());
		put(properties, CellUtil.TOP_BORDER_COLOR, style.getTopBorderColor());
		put(properties, CellUtil.WRAP_TEXT, style.getWrapText());
		return properties;
	}

	private static void setFormatProperties(CellStyle style, Workbook workbook, Map<String, Object> properties) {
		style.setAlignment(getHorizontalAlignment(properties, CellUtil.ALIGNMENT));
		style.setVerticalAlignment(getVerticalAlignment(properties, CellUtil.VERTICAL_ALIGNMENT));
		style.setBorderBottom(getBorderStyle(properties, CellUtil.BORDER_BOTTOM));
		style.setBorderLeft(getBorderStyle(properties, CellUtil.BORDER_LEFT));
		style.setBorderRight(getBorderStyle(properties, CellUtil.BORDER_RIGHT));
		style.setBorderTop(getBorderStyle(properties, CellUtil.BORDER_TOP));
		style.setBottomBorderColor(getShort(properties, CellUtil.BOTTOM_BORDER_COLOR));
		style.setDataFormat(getShort(properties, CellUtil.DATA_FORMAT));
		style.setFillPattern(getFillPattern(properties, CellUtil.FILL_PATTERN));
		style.setFillForegroundColor(getShort(properties, CellUtil.FILL_FOREGROUND_COLOR));
		style.setFillBackgroundColor(getShort(properties, CellUtil.FILL_BACKGROUND_COLOR));
		style.setFont(workbook.getFontAt(getInt(properties, CellUtil.FONT)));
		style.setHidden(getBoolean(properties, CellUtil.HIDDEN));
		style.setIndention(getShort(properties, CellUtil.INDENTION));
		style.setLeftBorderColor(getShort(properties, CellUtil.LEFT_BORDER_COLOR));
		style.setLocked(getBoolean(properties, CellUtil.LOCKED));
		style.setRightBorderColor(getShort(properties, CellUtil.RIGHT_BORDER_COLOR));
		style.setRotation(getShort(properties, CellUtil.ROTATION));
		style.setTopBorderColor(getShort(properties, CellUtil.TOP_BORDER_COLOR));
		style.setWrapText(getBoolean(properties, CellUtil.WRAP_TEXT));
	}

	private static void setFontProperties(Font font, Map<String, Object> properties) {
		font.setCharSet(getByte(properties, CHAR_SET));
		font.setItalic(getBoolean(properties, ITALIC));
		font.setFontName(getString(properties, FONT_NAME));
		font.setTypeOffset(getByte(properties, TYPE_OFFSET));
		font.setUnderline(getByte(properties, UNDERLINE));
		font.setFontHeight(getShort(properties, FONT_HEIGHT));
		font.setStrikeout(getBoolean(properties, STRIKEOUT));
		font.setBold(getBoolean(properties, BOLD));
		font.setColor(getShort(properties, COLOR));
	}

	private static void put(Map<String, Object> properties, String name, Object value) {
		properties.put(name, value);
	}

	private static Map<String, Object> getFontProperties(Font font) {
		Map<String, Object> properties = new HashMap<>();
		put(properties, CHAR_SET, font.getCharSet());
		put(properties, ITALIC, font.getItalic());
		put(properties, FONT_NAME, font.getFontName());
		put(properties, TYPE_OFFSET, font.getTypeOffset());
		put(properties, UNDERLINE, font.getUnderline());
		put(properties, FONT_HEIGHT, font.getFontHeight());
		put(properties, STRIKEOUT, font.getStrikeout());
		put(properties, BOLD, font.getBold());
		put(properties, COLOR, font.getColor());
		return properties;
	}

	private static void putFontProperties(final Map<String, Object> src, Map<String, Object> dest) {
		for (final String key : src.keySet()) {
			if (fontValues.contains(key) && src.containsKey(key)) {
				dest.put(key, src.get(key));
			}
		}
	}

	private static void putStyleProperties(final Map<String, Object> src, Map<String, Object> dest) {
		for (final String key : src.keySet()) {
			if (shortValues.contains(key)) {
				dest.put(key, getShort(src, key));
			} else if (intValues.contains(key)) {
				dest.put(key, getInt(src, key));
			} else if (booleanValues.contains(key)) {
				dest.put(key, getBoolean(src, key));
			} else if (borderTypeValues.contains(key)) {
				dest.put(key, getBorderStyle(src, key));
			} else if (CellUtil.ALIGNMENT.equals(key)) {
				dest.put(key, getHorizontalAlignment(src, key));
			} else if (CellUtil.VERTICAL_ALIGNMENT.equals(key)) {
				dest.put(key, getVerticalAlignment(src, key));
			} else if (CellUtil.FILL_PATTERN.equals(key)) {
				dest.put(key, getFillPattern(src, key));
			} else {
				log.debug("Ignoring unrecognized CellUtil format properties key: {}", key);
			}
		}
	}

	public static Byte getByte(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		if (value instanceof Byte) {
			return (Byte) value;
		}
		return 0;
	}

	public static String getString(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		if (value instanceof String) {
			return String.valueOf(value);
		}
		return "Arial";
	}

	public static boolean getBoolean(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		if (value instanceof Boolean) {
			return (Boolean) value;
		}
		return false;
	}

	public static short getShort(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		if (value instanceof Number) {
			return ((Number) value).shortValue();
		}
		return 0;
	}

	public static int getInt(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		if (value instanceof Number) {
			return ((Number) value).intValue();
		}
		return 0;
	}

	public static FillPatternType getFillPattern(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		FillPatternType pattern;
		if (value instanceof FillPatternType) {
			pattern = (FillPatternType) value;
		}
		// @deprecated 3.15 beta 2. getFillPattern will only work on FillPatternType enums instead of codes in the future.
		else if (value instanceof Short) {
			short code = (Short) value;
			pattern = FillPatternType.forInt(code);
		} else if (value == null) {
			pattern = FillPatternType.NO_FILL;
		} else {
			throw new RuntimeException("Unexpected fill pattern style class. Must be FillPatternType or Short (deprecated).");
		}
		return pattern;
	}

	public static BorderStyle getBorderStyle(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		BorderStyle border;
		if (value instanceof BorderStyle) {
			border = (BorderStyle) value;
		}
		// @deprecated 3.15 beta 2. getBorderStyle will only work on BorderStyle enums instead of codes in the future.
		else if (value instanceof Short) {
			short code = (Short) value;
			border = BorderStyle.valueOf(code);
		} else if (value == null) {
			border = BorderStyle.NONE;
		} else {
			throw new RuntimeException("Unexpected border style class. Must be BorderStyle or Short (deprecated).");
		}
		return border;
	}

	public static HorizontalAlignment getHorizontalAlignment(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		HorizontalAlignment align;
		if (value instanceof HorizontalAlignment) {
			align = (HorizontalAlignment) value;
		}
		// @deprecated 3.15 beta 2. getHorizontalAlignment will only work on HorizontalAlignment enums instead of codes in the future.
		else if (value instanceof Short) {
			short code = (Short) value;
			align = HorizontalAlignment.forInt(code);
		} else if (value == null) {
			align = HorizontalAlignment.GENERAL;
		} else {
			throw new RuntimeException("Unexpected horizontal alignment style class. Must be HorizontalAlignment or Short (deprecated).");
		}
		return align;
	}

	public static VerticalAlignment getVerticalAlignment(Map<String, Object> properties, String name) {
		Object value = properties.get(name);
		VerticalAlignment align;
		if (value instanceof VerticalAlignment) {
			align = (VerticalAlignment) value;
		}
		// @deprecated 3.15 beta 2. getVerticalAlignment will only work on VerticalAlignment enums instead of codes in the future.
		else if (value instanceof Short) {
			short code = (Short) value;
			align = VerticalAlignment.forInt(code);
		} else if (value == null) {
			align = VerticalAlignment.BOTTOM;
		} else {
			throw new RuntimeException("Unexpected vertical alignment style class. Must be VerticalAlignment or Short (deprecated).");
		}
		return align;
	}

	public static void setCellValue(Cell cell, Object value) {
		if (value == null) {
			cell.setBlank();
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Number) {
			double doubleValue = ((Number) value).doubleValue();
			cell.setCellValue(doubleValue);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);
		} else if (value instanceof RichTextString) {
			cell.setCellValue((RichTextString) value);
		} else if (value instanceof Hyperlink) {
			Hyperlink hyperlink = (Hyperlink) value;
			cell.setHyperlink(hyperlink);
			cell.setCellValue(hyperlink.getLabel());
		} else if (isFormulaDefinition(value)) {
			cell.setCellFormula(((String) value).substring(1));
		} else {
			cell.setCellValue(value.toString());
		}
	}

	private static boolean isFormulaDefinition(Object obj) {
		if (obj instanceof String) {
			String str = (String) obj;
			return str.length() >= 2 && str.charAt(0) == '=';
		} else {
			return false;
		}
	}

	public static String getCellValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		final String result;

		switch (cell.getCellType()) {
			case BLANK:
			case _NONE:
				result = null;
				break;
			case BOOLEAN:
				result = Boolean.toString(cell.getBooleanCellValue());
				break;
			case ERROR:
				result = getErrorResult(cell);
				break;
			case FORMULA:
				Workbook workbook = cell.getRow().getSheet().getWorkbook();
				result = getFormulaCellValue(workbook, cell);
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					result = date == null ? null : DateFormatUtils.format(date, DATE_TIME_FORMAT);
				} else {
					result = getNumericCellValueAsString(cell.getCellStyle(), cell.getNumericCellValue());
				}
				break;
			case STRING:
				result = cell.getRichStringCellValue().getString();
				break;
			default:
				throw new IllegalStateException("Unknown cell type: " + cell.getCellType());
		}
		return result;
	}

	private static Object getCellValueAsObject(final Workbook workbook, final Cell cell) {
		if (cell == null) {
			return null;
		}

		final Object result;

		switch (cell.getCellType()) {
			case BLANK:
			case _NONE:
				result = null;
				break;
			case BOOLEAN:
				result = cell.getBooleanCellValue();
				break;
			case ERROR:
				result = getErrorResult(cell);
				break;
			case FORMULA:
				result = getFormulaCellValueAsObject(workbook, cell);
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					result = cell.getDateCellValue();
				} else {
					result = getDoubleAsNumber(cell.getNumericCellValue());
				}
				break;
			case STRING:
				result = cell.getRichStringCellValue().getString();
				break;
			default:
				throw new IllegalStateException("Unknown cell type: " + cell.getCellType());
		}
		return result;
	}

	private static String getErrorResult(final Cell cell) {
		try {
			return FormulaError.forInt(cell.getErrorCellValue()).getString();
		} catch (final RuntimeException e) {
			log.debug("Getting error code for ({},{}) failed!: {}", cell.getRowIndex(), cell.getColumnIndex(), e.getMessage());
			if (cell instanceof XSSFCell) {
				return ((XSSFCell) cell).getErrorCellString();
			} else {
				log.error("Couldn't handle unexpected error scenario in cell: ({},{})", cell.getRowIndex(), cell.getColumnIndex());
				throw e;
			}
		}
	}

	private static String getFormulaCellValue(Workbook workbook, Cell cell) {
		try {
			double numericCellValue = cell.getNumericCellValue();
			return getNumericCellValueAsString(cell.getCellStyle(), numericCellValue);
		} catch (Exception e) {
			log.warn("Failed to fetch cached formula value of cell: " + cell, e);
		}

		final Cell evaluatedCell = getEvaluatedCell(workbook, cell);
		if (evaluatedCell != null) {
			return getCellValue(evaluatedCell);
		} else {
			return cell.getCellFormula();
		}
	}

	private static Object getFormulaCellValueAsObject(final Workbook workbook, final Cell cell) {
		try {
			return getDoubleAsNumber(cell.getNumericCellValue());
		} catch (final Exception e) {
			log.warn("Failed to fetch cached formula value of cell: " + cell, e);
		}

		// evaluate cell first, if possible
		final Cell evaluatedCell = getEvaluatedCell(workbook, cell);
		if (evaluatedCell != null) {
			return getCellValueAsObject(workbook, evaluatedCell);
		} else {
			return cell.getCellFormula();
		}
	}

	private static Cell getEvaluatedCell(final Workbook workbook, final Cell cell) {
		try {
			final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			return evaluator.evaluateInCell(cell);
		} catch (RuntimeException e) {
			log.warn("Exception occurred while evaluating formula at position ({},{}): {}", cell.getRowIndex(),
				cell.getColumnIndex(), e.getMessage());
		}
		return null;
	}

	private static Number getDoubleAsNumber(final double value) {
		if (value % 1 == 0 && value <= Integer.MAX_VALUE) {
			return (int) value;
		} else {
			return value;
		}
	}

	private static String getNumericCellValueAsString(final CellStyle cellStyle, final double cellValue) {
		final int formatIndex = cellStyle.getDataFormat();
		String formatString = cellStyle.getDataFormatString();
		if (formatString == null) {
			formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
		}
		final DataFormatter formatter = new DataFormatter();
		return formatter.formatRawCellContents(cellValue, formatIndex, formatString);
	}
}
