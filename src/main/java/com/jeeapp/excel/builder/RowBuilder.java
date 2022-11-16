package com.jeeapp.excel.builder;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.RegExUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.PropertyAccessorFactory;
import org.springframework.core.ResolvableType;
import org.springframework.core.annotation.AnnotationUtils;
import org.springframework.util.ObjectUtils;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.model.Cell;
import com.jeeapp.excel.model.Column;

/**
 * @author justice
 */
public class RowBuilder<T> {

	public static final Map<Class<?>, List<Field>> FIELDS_CACHE = new ConcurrentHashMap<>();

	private final Class<T> type;

	private final SheetBuilder parent;

	private final List<String> properties;

	private int thisRow;

	private int lastRow;

	private int thisCol = -1;

	public RowBuilder(SheetBuilder parent, Class<T> type) {
		Validate.notNull(type, "Type must not be null");
		this.parent = parent;
		this.type = type;
		this.properties = getProperties(null, type);
	}

	/**
	 * 创建表头
	 */
	private void createHeader(Column header) {
		int firstRow = header.getFirstRow();
		int lastRow = header.getFirstRow();
		int firstCol = header.getFirstCol();
		int lastCol = header.getLastCol();
		if (CollectionUtils.isNotEmpty(header.getChildren())) {
			for (Column child : header.getChildren()) {
				createHeader(child);
			}
		} else {
			lastRow = this.lastRow;
		}

		// 设置表头样式
		int width = header.getWidth() == 8 ? parent.sheet.getDefaultColumnWidth() : header.getWidth();
		parent.setColumnWidth(firstCol, width)
			.matchingColumn(firstCol)
			.setDataFormat(header.getFormat())
			.end()
			.matchingRegion(this.lastRow + 1, parent.maxRows - this.lastRow - 1, lastCol, lastCol)
			.createConstraint(header.getValidationType(),
				header.getOperatorType(),
				header.getFirstFormula(),
				header.getSecondFormula(),
				header.getExplicitListValues(),
				header.getDateFormat())
			.allowedEmptyCell(header.isAllowEmpty())
			.setErrorStyle(header.getErrorStyle())
			.showErrorBox(header.isShowErrorBox(), header.getErrorBoxTitle(), header.getErrorBoxText())
			.showPromptBox(header.isShowPromptBox(), header.getPromptBoxTitle(), header.getPromptBoxText())
			.end();
		if (firstRow == lastRow && firstCol == lastCol) {
			parent.matchingCell(firstRow, firstCol)
				.setFillForegroundColor(header.getFillForegroundColor())
				.setFillBackgroundColor(header.getFillBackgroundColor())
				.setFillPattern(header.getFillPatternType())
				.setFontColor(header.getFontColor())
				.setFontBold(true)
				.setBorder(header.getBorder())
				.setBorderColor(header.getBorderColor())
				.setCommentText(header.getComment())
				.setCommentAuthor(header.getCommentAuthor())
				.setCommentSize(header.getCommentWidth(), header.getCommentHeight())
				.setCellValue(header.getValue());
		} else {
			parent.matchingRegion(firstRow, lastRow, firstCol, lastCol)
				.setFillForegroundColor(header.getFillForegroundColor())
				.setFillBackgroundColor(header.getFillBackgroundColor())
				.setFillPattern(header.getFillPatternType())
				.setFontColor(header.getFontColor())
				.setFontBold(true)
				.setBorder(header.getBorder())
				.setBorderColor(header.getBorderColor())
				.addMergedRegion()
				.setCommentText(header.getComment())
				.setCommentAuthor(header.getCommentAuthor())
				.setCommentSize(header.getCommentWidth(), header.getCommentHeight())
				.setCellValue(header.getValue());
		}
	}

	public RowBuilder<T> createRow(T object) {
		Validate.notNull(object, "object must be not null");
		thisRow = parent.sheet.getLastRowNum() + 1;
		lastRow = thisRow;
		List<Cell> cells = resolveCells(object);
		for (Cell cell : cells) {
			int firstRow = cell.getFirstRow();
			int lastRow = cell.getLastRow();
			int firstCol = cell.getFirstCol();
			int lastCol = cell.getLastCol();
			if (firstRow == lastRow && firstCol == lastCol) {
				parent.createCell(firstRow, firstCol, cell.getValue());
			} else {
				parent.createCell(firstRow, firstCol, cell.getValue())
					.matchingRegion(firstRow, lastRow, firstCol, lastCol)
					.addMergedRegion()
					.end();
			}
		}
		return this;
	}

	/**
	 * 创建空行
	 */
	public RowBuilder<T> createRow() {
		parent.createRow();
		return this;
	}

	/**
	 * 创建单行
	 */
	public RowBuilder<T> createRow(Object... cells) {
		parent.createRow(cells);
		return this;
	}

	/**
	 * 创建多行
	 */
	public RowBuilder<T> createRows(Object[][] rows) {
		parent.createRows(rows);
		return this;
	}

	/**
	 * 当前行指定列创建批注
	 */
	@Deprecated
	public RowBuilder<T> createCellComment(String comment, String author, int col1, int row2, int col2) {
		parent.createCellComment(comment, author, thisRow, col1, row2, col2);
		return this;
	}

	/**
	 * 对象行
	 */
	public RowBuilder<T> createRows(Collection<T> beans) {
		if (CollectionUtils.isNotEmpty(beans)) {
			for (T bean : beans) {
				createRow(bean);
			}
		}
		return this;
	}

	/**
	 * 创建表头
	 */
	public RowBuilder<T> createHeader(String... names) {
		if (ArrayUtils.isNotEmpty(names)) {
			// 只保留指定属性，包含子属性
			properties.removeIf(property -> !matchesProperty(Arrays.asList(names), property));
		}
		thisRow = parent.sheet.getLastRowNum() + 1;
		lastRow = thisRow;
		List<Column> headers = resolveHeaders(null, type);
		for (Column header : headers) {
			createHeader(header);
		}
		return this;
	}

	public SheetBuilder end() {
		return parent.end();
	}

	public Workbook build() {
		return parent.build();
	}

	/**
	 * 判断属性是否需要过滤
	 */
	private Boolean matchesProperty(List<String> properties, String property) {
		return properties.contains(property)
			|| IterableUtils.matchesAny(properties, it -> it.startsWith(property + "."));
	}

	/**
	 * 创建表头
	 */
	protected List<Column> resolveHeaders(Column column, Class<?> type) {
		List<Column> headers = new ArrayList<>();
		for (Field field : getSortedFields(type)) {
			Class<?> fieldType = field.getType();
			String property = column == null ? field.getName() : column.getName() + "." + field.getName();
			if (!matchesProperty(properties, property)) {
				continue;
			}
			Column header = Column.of(property, field);
			// 根据嵌套属性中的点来获取开始行位置，每次递归都会创建一行
			header.setFirstRow(StringUtils.countMatches(property, ".") + thisRow);
			// 更新表头最后一行的位置，用于表头合并
			lastRow = Math.max(header.getFirstRow(), lastRow);
			// 更新表头列的位置，每个属性对应一列
			header.setFirstCol(thisCol + 1);
			// 如果存在嵌套属性，根据嵌套属性创建子表头
			if (!BeanUtils.isSimpleProperty(fieldType)) {
				if (fieldType.isArray()) {
					header.setChildren(resolveHeaders(header, fieldType.getComponentType()));
				} else if (Collection.class.isAssignableFrom(fieldType)) {
					header.setChildren(resolveHeaders(header, ResolvableType.forField(field).resolveGeneric(0)));
				} else {
					header.setChildren(resolveHeaders(header, fieldType));
				}
			}
			if (CollectionUtils.isEmpty(header.getChildren())) {
				thisCol++;
			} else {
				properties.remove(property);
			}
			header.setLastCol(thisCol);
			headers.add(header);
		}
		return headers;
	}

	/**
	 * 根据目标对象创建单元格
	 */
	protected List<Cell> resolveCells(T target) {
		List<String> nestedPropertyPaths = new ArrayList<>();
		Map<String, Integer> nestedPropertyPathRowCount = new HashMap<>();
		List<Cell> cells = resolveCells(target, StringUtils.EMPTY, nestedPropertyPaths, nestedPropertyPathRowCount);

		// 计算属性值行距
		Map<String, RowSpan> propertyPathRowSpans = new HashMap<>();
		for (String propertyPath : nestedPropertyPaths) {
			int rowSpans = nestedPropertyPathRowCount.get(propertyPath) == null ? 0 : nestedPropertyPathRowCount.get(propertyPath) - 1;
			String[] indexes = StringUtils.substringsBetween(propertyPath, "[", "]");
			int firstRow = Integer.parseInt(indexes[indexes.length - 1]);
			if (firstRow == 0) {
				String parentPropertyPath = StringUtils.substringBeforeLast(propertyPath, ".");
				if (propertyPathRowSpans.get(parentPropertyPath) != null) {
					firstRow = propertyPathRowSpans.get(parentPropertyPath).getFirstRow() + firstRow;
				}
			} else {
				String parentPropertyPath = StringUtils.substringBeforeLast(propertyPath, "[") + "[" + (firstRow - 1) + "]";
				firstRow = propertyPathRowSpans.get(parentPropertyPath).getLastRow() + 1;
			}
			int lastRow = firstRow + rowSpans;
			propertyPathRowSpans.put(propertyPath, new RowSpan(firstRow, lastRow));
			this.lastRow = Math.max(this.lastRow, lastRow + thisRow);
		}

		// 设置单元格行距
		for (Cell cell : cells) {
			if (cell.getName().contains(".")) {
				String propertyPath = StringUtils.substringBeforeLast(cell.getName(), ".");
				cell.setFirstRow(propertyPathRowSpans.get(propertyPath).getFirstRow() + thisRow);
				cell.setLastRow(propertyPathRowSpans.get(propertyPath).getLastRow() + thisRow);
			} else {
				cell.setFirstRow(thisRow);
				cell.setLastRow(lastRow);
			}
		}
		return cells;
	}

	/**
	 * 转换成单元格
	 */
	private List<Cell> resolveCells(Object target, String parentPropertyPath, List<String> nestedPropertyPaths,
		Map<String, Integer> nestedPropertyPathRowCount) {
		List<Cell> cells = new ArrayList<>();
		BeanWrapper beanWrapper = PropertyAccessorFactory.forBeanPropertyAccess(target);
		for (Field field : getSortedFields(beanWrapper.getWrappedClass())) {
			Class<?> propertyType = field.getType();
			String propertyName = parentPropertyPath + field.getName();
			String property = RegExUtils.removeAll(propertyName, "\\[(.*?)]");
			int column = properties.indexOf(property);
			if (IterableUtils.matchesAny(properties, str -> str.startsWith(property + ".") || str.equals(property))) {
				Object propertyValue = beanWrapper.getPropertyValue(field.getName());
				if (propertyType.isArray() && propertyValue != null) {
					Object[] values = ObjectUtils.toObjectArray(propertyValue);
					int length = values.length;
					calculateNestedPropertyPathRowCount(parentPropertyPath, length, nestedPropertyPathRowCount);
					for (int i = 0; i < length; i++) {
						String propertyPath = propertyName + "[" + i + "].";
						nestedPropertyPaths.add(StringUtils.substringBeforeLast(propertyPath, "."));
						cells.addAll(resolveCells(values[i], propertyPath, nestedPropertyPaths, nestedPropertyPathRowCount));
					}
				} else if (Collection.class.isAssignableFrom(propertyType) && propertyValue != null) {
					Collection<?> values = (Collection<?>) propertyValue;
					int size = values.size();
					calculateNestedPropertyPathRowCount(parentPropertyPath, size, nestedPropertyPathRowCount);
					for (int i = 0; i < size; i++) {
						String propertyPath = propertyName + "[" + i + "].";
						nestedPropertyPaths.add(StringUtils.substringBeforeLast(propertyPath, "."));
						cells.addAll(resolveCells(values.toArray()[i], propertyPath, nestedPropertyPaths, nestedPropertyPathRowCount));
					}
				} else if (!BeanUtils.isSimpleProperty(propertyType) && propertyValue != null) {
					String propertyPath = propertyName + ".";
					nestedPropertyPaths.add(StringUtils.substringBeforeLast(propertyPath, "."));
					cells.addAll(resolveCells(propertyValue, propertyPath, nestedPropertyPaths, nestedPropertyPathRowCount));
				} else if (column > -1) {
					Cell cell = Cell.of(propertyName, propertyValue);
					cell.setFirstCol(column);
					cell.setLastCol(column);
					cells.add(cell);
				}
			}
		}
		return cells;
	}

	/**
	 * 计算属性值占用行数
	 */
	private void calculateNestedPropertyPathRowCount(String parentPropertyPath, int total,
		Map<String, Integer> nestedPropertyPathRowCount) {
		nestedPropertyPathRowCount.merge(StringUtils.substringBeforeLast(parentPropertyPath, "."), total, Math::max);
		int length = StringUtils.split(parentPropertyPath, ".").length;
		if (length == 0) {
			return;
		}
		for (int i = 1; i < length; i++) {
			String propertyPath = StringUtils.substring(parentPropertyPath, 0, StringUtils.ordinalIndexOf(parentPropertyPath, ".", i));
			nestedPropertyPathRowCount.put(propertyPath, nestedPropertyPathRowCount.get(propertyPath) + total - 1);
		}
	}

	/**
	 * 获取所有属性
	 */
	private static List<String> getProperties(String parentProperty, Class<?> type) {
		List<String> properties = new ArrayList<>();
		for (Field field : getSortedFields(type)) {
			Class<?> fieldType = field.getType();
			String property = parentProperty == null ? field.getName() : parentProperty + "." + field.getName();
			if (!BeanUtils.isSimpleProperty(fieldType)) {
				if (fieldType.isArray()) {
					properties.addAll(getProperties(property, fieldType.getComponentType()));
				} else if (Collection.class.isAssignableFrom(fieldType)) {
					properties.addAll(getProperties(property, ResolvableType.forField(field).resolveGeneric(0)));
				} else {
					properties.addAll(getProperties(property, fieldType));
				}
			}
			properties.add(property);
		}
		return properties;
	}

	/**
	 * 获取字段
	 */
	private static List<Field> getSortedFields(Class<?> type) {
		return FIELDS_CACHE.computeIfAbsent(type, key -> FieldUtils.getAllFieldsList(type)
			.stream()
			.filter(field -> {
				// 只获取带有@ExcelProperty注解的字段
				ExcelProperty annotation = AnnotationUtils.getAnnotation(field, ExcelProperty.class);
				if (annotation == null) {
					return false;
				}
				if (Modifier.isStatic(field.getModifiers()) && Modifier.isFinal(field.getModifiers())) {
					return false;
				}
				return !Modifier.isTransient(field.getModifiers());
			}).sorted(Comparator.comparingInt(field -> {
				ExcelProperty annotation = AnnotationUtils.getAnnotation(field, ExcelProperty.class);
				if (annotation == null) {
					return Integer.MAX_VALUE;
				}
				return annotation.column();
			})).collect(Collectors.toList()));
	}

	@Data
	@AllArgsConstructor
	static class RowSpan {

		private int firstRow;

		private int lastRow;
	}
}
