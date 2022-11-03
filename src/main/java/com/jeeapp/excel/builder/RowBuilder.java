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
import org.apache.poi.ss.util.CellRangeAddressList;
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
		// 表头样式
		int defaultWidth = parent.sheet.getDefaultColumnWidth();
		int width = header.getWidth() == 8 ? defaultWidth : header.getWidth();
		parent.setColumnWidth(firstCol, width)
			.matchingColumn(firstCol)
			.setDataFormat(header.getFormat())
			.addCellStyle()
			.matchingRegion(firstRow, lastRow, firstCol, lastCol)
			.setFillForegroundColor(header.getFillForegroundColor())
			.setFillBackgroundColor(header.getFillBackgroundColor())
			.setFillPattern(header.getFillPatternType())
			.setFontColor(header.getFontColor())
			.setFontBold(true)
			.addCellStyle()
			.createCell(firstRow, firstCol, header.getValue());

		// 数据验证
		if (header.hasValidation()) {
			parent.addValidationData(
				new CellRangeAddressList(firstRow, parent.maxRows - firstRow - 1, lastCol, lastCol),
				header.getValidationType(),
				header.getOperatorType(),
				header.getFirstFormula(),
				header.getSecondFormula(),
				header.getExplicitListValues(),
				header.isAllowEmpty(),
				header.getErrorStyle(),
				header.isShowPromptBox(),
				header.getPromptBoxTitle(),
				header.getPromptBoxText(),
				header.isShowErrorBox(),
				header.getErrorBoxTitle(),
				header.getErrorBoxText());
		}

		// 表头标注
		if (header.hasComment()) {
			parent.createCellComment(header.getComment(),
				header.getCommentAuthor(),
				header.getFirstRow(),
				header.getFirstCol(),
				header.getCommentHeight(),
				header.getCommentWidth());
		}

		if (CollectionUtils.isNotEmpty(header.getChildren())) {
			for (Column child : header.getChildren()) {
				createHeader(child);
			}
		} else {
			lastRow = parent.lastRow;
		}
		// 合并表头样式
		parent.addCellRange(firstRow, lastRow, firstCol, lastCol)
			.setBorder(header.getBorder())
			.setBorderColor(header.getBorderColor())
			.merge();
	}

	public RowBuilder<T> createRow(T bean) {
		Validate.notNull(bean, "bean must be not null");
		thisRow = parent.lastRow + 1;
		parent.lastRow = thisRow;
		List<Cell> cells = createCells(bean);
		for (Cell cell : cells) {
			int firstRow = cell.getFirstRow();
			int lastRow = cell.getLastRow();
			int firstCol = cell.getFirstCol();
			int lastCol = cell.getLastCol();
			parent.createCell(firstRow, firstCol, cell.getValue())
				.addCellRange(firstRow, lastRow, firstCol, lastCol)
				.merge();
		}
		// 补全空白边框
		if (parent.lastCol > -1) {
			parent.addCellRange(thisRow, parent.lastRow, 0, parent.lastCol);
		}
		return this;
	}

	/**
	 * 创建单行
	 */
	public RowBuilder<T> createRow(Object[] cells) {
		parent.createRow(cells);
		thisRow = parent.lastRow;
		return this;
	}

	/**
	 * 创建多行
	 */
	public RowBuilder<T> createRows(Object[][] rows) {
		parent.createRows(rows);
		thisRow = parent.lastRow;
		return this;
	}

	/**
	 * 当前行指定列创建批注
	 */
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
		this.thisRow = parent.lastRow + 1;
		List<Column> headers = createHeader(null, type);
		for (Column header : headers) {
			createHeader(header);
		}
		return this;
	}

	/**
	 * 写入结束
	 */
	public SheetBuilder end() {
		return parent;
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
	protected List<Column> createHeader(Column column, Class<?> type) {
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
			parent.lastRow = Math.max(header.getFirstRow(), parent.lastRow);
			// 更新表头列的位置，每个属性对应一列
			header.setFirstCol(parent.lastCol + 1);
			// 如果存在嵌套属性，根据嵌套属性创建子表头
			if (!BeanUtils.isSimpleProperty(fieldType)) {
				if (fieldType.isArray()) {
					header.setChildren(createHeader(header, fieldType.getComponentType()));
				} else if (Collection.class.isAssignableFrom(fieldType)) {
					header.setChildren(createHeader(header, ResolvableType.forField(field).resolveGeneric(0)));
				} else {
					header.setChildren(createHeader(header, fieldType));
				}
			}
			if (CollectionUtils.isEmpty(header.getChildren())) {
				parent.lastCol++;
			} else {
				properties.remove(property);
			}
			header.setLastCol(parent.lastCol);
			headers.add(header);
		}
		return headers;
	}

	/**
	 * 根据目标对象创建单元格
	 */
	protected List<Cell> createCells(T target) {
		List<String> nestedPropertyPaths = new ArrayList<>();
		Map<String, Integer> nestedPropertyPathRowCount = new HashMap<>();
		List<Cell> cells = getCells(target, StringUtils.EMPTY, nestedPropertyPaths, nestedPropertyPathRowCount);

		// 计算属性值行距
		Map<String, RowSpan> propertyPathRowSpans = new HashMap<>();
		for (String propertyPath : nestedPropertyPaths) {
			long rowSpans = nestedPropertyPathRowCount.get(propertyPath) == null ? 0 : nestedPropertyPathRowCount.get(propertyPath) - 1;
			String[] indexes = StringUtils.substringsBetween(propertyPath, "[", "]");
			long firstRow = Integer.parseInt(indexes[indexes.length - 1]);
			if (firstRow == 0) {
				String parentPropertyPath = StringUtils.substringBeforeLast(propertyPath, ".");
				if (propertyPathRowSpans.get(parentPropertyPath) != null) {
					firstRow = propertyPathRowSpans.get(parentPropertyPath).getFirstRow() + firstRow;
				}
			} else {
				String parentPropertyPath = StringUtils.substringBeforeLast(propertyPath, "[") + "[" + (firstRow - 1) + "]";
				firstRow = propertyPathRowSpans.get(parentPropertyPath).getLastRow() + 1;
			}
			long lastRow = firstRow + rowSpans;
			propertyPathRowSpans.put(propertyPath, new RowSpan((int) firstRow, (int) lastRow));
			parent.lastRow = Math.max(parent.lastRow, (int) lastRow + thisRow);
		}

		// 设置单元格行距
		for (Cell cell : cells) {
			if (cell.getName().contains(".")) {
				String propertyPath = StringUtils.substringBeforeLast(cell.getName(), ".");
				cell.setFirstRow(propertyPathRowSpans.get(propertyPath).getFirstRow() + thisRow);
				cell.setLastRow(propertyPathRowSpans.get(propertyPath).getLastRow() + thisRow);
			} else {
				cell.setFirstRow(thisRow);
				cell.setLastRow(parent.lastRow);
			}
		}
		return cells;
	}

	/**
	 * 转换成单元格
	 */
	private List<Cell> getCells(Object target, String parentPropertyPath, List<String> nestedPropertyPaths,
		Map<String, Integer> nestedPropertyPathRowCount) {
		List<Cell> cells = new ArrayList<>();
		BeanWrapper beanWrapper = PropertyAccessorFactory.forBeanPropertyAccess(target);
		for (Field field : getSortedFields(beanWrapper.getWrappedClass())) {
			Class<?> propertyType = field.getType();
			String propertyName = parentPropertyPath + field.getName();
			String property = RegExUtils.removeAll(propertyName, "\\[(.*?)]");
			int col = properties.indexOf(property);
			if (IterableUtils.matchesAny(properties, str -> str.startsWith(property + ".") || str.equals(property))) {
				Object propertyValue = beanWrapper.getPropertyValue(field.getName());
				if (propertyType.isArray() && propertyValue != null) {
					Object[] values = ObjectUtils.toObjectArray(propertyValue);
					int length = values.length;
					calculateNestedPropertyPathRowCount(parentPropertyPath, length, nestedPropertyPathRowCount);
					for (int i = 0; i < length; i++) {
						String propertyPath = propertyName + "[" + i + "].";
						nestedPropertyPaths.add(StringUtils.substringBeforeLast(propertyPath, "."));
						cells.addAll(getCells(values[i], propertyPath, nestedPropertyPaths, nestedPropertyPathRowCount));
					}
				} else if (Collection.class.isAssignableFrom(propertyType) && propertyValue != null) {
					Collection<?> values = (Collection<?>) propertyValue;
					int size = values.size();
					calculateNestedPropertyPathRowCount(parentPropertyPath, size, nestedPropertyPathRowCount);
					for (int i = 0; i < size; i++) {
						String propertyPath = propertyName + "[" + i + "].";
						nestedPropertyPaths.add(StringUtils.substringBeforeLast(propertyPath, "."));
						cells.addAll(getCells(values.toArray()[i], propertyPath, nestedPropertyPaths, nestedPropertyPathRowCount));
					}
				} else if (!BeanUtils.isSimpleProperty(propertyType) && propertyValue != null) {
					String propertyPath = propertyName + ".";
					nestedPropertyPaths.add(StringUtils.substringBeforeLast(propertyPath, "."));
					cells.addAll(getCells(propertyValue, propertyPath, nestedPropertyPaths, nestedPropertyPathRowCount));
				} else if (col > -1) {
					Cell cell = Cell.of(propertyName, propertyValue);
					cell.setFirstCol(col);
					cell.setLastCol(col);
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
