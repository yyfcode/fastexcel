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
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.RegExUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.PropertyAccessorFactory;
import org.springframework.core.ResolvableType;
import org.springframework.core.annotation.AnnotationUtils;
import org.springframework.util.Assert;
import org.springframework.util.ObjectUtils;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.model.Cell;
import com.jeeapp.excel.model.Header;

/**
 * @author justice
 */
@Slf4j
public class TableBuilder<T> extends SheetBuilderHelper {

	public static final Map<Class<?>, List<Field>> FIELDS_CACHE = new ConcurrentHashMap<>();

	public static final Map<Class<?>, List<String>> PROPERTIES_CACHE = new ConcurrentHashMap<>();

	private final Class<T> type;

	private final SheetBuilder parent;

	private final List<String> properties;

	private int thisRow;

	private int lastRow;

	private int thisCol = -1;

	private int lastCol;

	protected TableBuilder(SheetBuilder parent, Class<T> type) {
		super(parent, parent.sheet);
		this.parent = parent;
		this.type = type;
		this.properties = getProperties(type);
		this.lastCol = properties.size() - 1;
	}

	/**
	 * 创建表头
	 */
	private void createHeader(Header header) {
		int firstRow = header.getFirstRow();
		int lastRow = header.getFirstRow();
		int firstCol = header.getFirstCol();
		int lastCol = header.getLastCol();
		if (CollectionUtils.isNotEmpty(header.getChildren())) {
			for (Header child : header.getChildren()) {
				createHeader(child);
			}
		} else {
			lastRow = this.lastRow;
		}

		// column style
		parent.matchingColumn(firstCol)
			.setColumnWidth(header.getWidth())
			.setColumnHidden(header.getHidden())
			.setDataFormat(header.getFormat())
			.addCellStyle();
		// column validation
		if (header.getValidationType() > ValidationType.ANY && header.getValidationType() <= ValidationType.FORMULA) {
			parent.matchingRegion(this.lastRow + 1, parent.maxRows - this.lastRow - 1, lastCol, lastCol)
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
				.addValidationData();
		}
		// header comment
		if (StringUtils.isNotBlank(header.getComment())) {
			parent.matchingCell(firstRow, firstCol)
				.createCellComment(header.getComment(),
					header.getCommentAuthor(),
					header.getCommentWidth(),
					header.getCommentHeight());
		}
		// header style
		if (firstRow == lastRow && firstCol == lastCol) {
			parent.matchingCell(firstRow, firstCol)
				.setFillForegroundColor(header.getFillForegroundColor())
				.setFillBackgroundColor(header.getFillBackgroundColor())
				.setFillPattern(header.getFillPatternType())
				.setFontColor(header.getFontColor())
				.setFontBold(true)
				.setBorder(header.getBorder())
				.setBorderColor(header.getBorderColor())
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
				.mergeRegion()
				.setCellValue(header.getValue());
		}
	}

	public TableBuilder<T> createRow(T object) {
		Assert.notNull(object, "object must be not null");
		thisRow = parent.sheet.getLastRowNum() + 1;
		lastRow = thisRow;
		List<Cell> cells = resolveCells(object);
		int maxRow = 0;
		for (Cell cell : cells) {
			int firstRow = cell.getFirstRow();
			int lastRow = cell.getLastRow();
			int firstCol = cell.getFirstCol();
			int lastCol = cell.getLastCol();
			if (firstRow == lastRow && firstCol == lastCol) {
				parent.createCell(firstRow, firstCol, cell.getValue());
			} else {
				parent.matchingRegion(firstRow, lastRow, firstCol, lastCol)
					.mergeRegion()
					.setCellValue(cell.getValue());
			}
			maxRow = Math.max(lastRow, maxRow);
		}
		// 填充未定义的单元
		parent.matchingRegion(thisRow, maxRow, 0, lastCol).fillUndefinedCells();
		return this;
	}

	/**
	 * 当前行指定列创建批注
	 * @deprecated removed in 0.1.0, use {@link SheetBuilder#matchingCell(CellAddress)} instead.
	 */
	@Deprecated
	public TableBuilder<T> createCellComment(String comment, String author, int col1, int row2, int col2) {
		parent.createCellComment(comment, author, thisRow, col1, row2, col2);
		return this;
	}

	/**
	 * 对象行
	 */
	public TableBuilder<T> createRows(Collection<T> beans) {
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
	public TableBuilder<T> createHeader(String... names) {
		if (ArrayUtils.isNotEmpty(names)) {
			properties.removeIf(property -> !IterableUtils.matchesAny(Arrays.asList(names),
				it -> it.equals(property) || property.startsWith(it + ".")));
		}
		thisRow = parent.sheet.getLastRowNum() + 1;
		lastRow = thisRow;
		lastCol = properties.size() - 1;
		List<Header> headers = resolveHeaders(null, type);
		for (Header header : headers) {
			createHeader(header);
		}
		return this;
	}

	public SheetBuilder createSheet() {
		return parent.createSheet();
	}

	public SheetBuilder createSheet(String sheetName) {
		return parent.createSheet(sheetName);
	}

	public WorkbookBuilder end() {
		return parent.end();
	}

	public Workbook build() {
		return parent.build();
	}

	@Override
	protected SheetBuilder self() {
		return parent;
	}

	/**
	 * 创建表头
	 */
	private List<Header> resolveHeaders(Header parent, Class<?> type) {
		List<Header> headers = new ArrayList<>();
		for (Field field : getFields(type)) {
			Class<?> fieldType = field.getType();
			String property = parent == null ? field.getName() : parent.getName() + "." + field.getName();
			if (!IterableUtils.matchesAny(properties, it -> it.equals(property) || it.startsWith(property + "."))) {
				continue;
			}
			Header header = new Header(field);
			header.setName(property);
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
			}
			header.setLastCol(thisCol);
			headers.add(header);
		}
		return headers;
	}

	/**
	 * 根据目标对象创建单元格
	 */
	private List<Cell> resolveCells(T target) {
		List<String> nestedPropertyPaths = new ArrayList<>();
		Map<String, Integer> nestedPropertyPathRowCount = new HashMap<>();
		List<Cell> cells = resolveCells(target, StringUtils.EMPTY, nestedPropertyPaths, nestedPropertyPathRowCount);

		// 计算属性值行距
		Map<String, RowSpan> propertyPathRowSpans = new HashMap<>();
		for (String propertyPath : nestedPropertyPaths) {
			Integer rowCount = nestedPropertyPathRowCount.get(propertyPath);
			int rowSpans = rowCount == null || rowCount == 0 ? 0 : rowCount - 1;
			String[] indexes = StringUtils.substringsBetween(propertyPath, "[", "]");
			int firstRow = indexes == null ? 0 : Integer.parseInt(indexes[indexes.length - 1]);
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
		for (Field field : getFields(beanWrapper.getWrappedClass())) {
			Class<?> propertyType = field.getType();
			String propertyName = parentPropertyPath + field.getName();
			String property = RegExUtils.removeAll(propertyName, "\\[(.*?)]");
			int column = properties.indexOf(property);
			if (!IterableUtils.matchesAny(properties, it -> it.equals(property) || it.startsWith(property + "."))) {
				continue;
			}
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
				Cell cell = new Cell(field);
				cell.setName(propertyName);
				cell.setValue(propertyValue);
				cell.setFirstCol(column);
				cell.setLastCol(column);
				cells.add(cell);
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
	 * 从缓存中获取所有属性
	 */
	private List<String> getProperties(Class<?> type) {
		return new ArrayList<>(PROPERTIES_CACHE.computeIfAbsent(type, key -> getProperties(null, type)));
	}

	/**
	 * 获取所有属性
	 */
	private List<String> getProperties(String parentProperty, Class<?> type) {
		List<String> properties = new ArrayList<>();
		for (Field field : getFields(type)) {
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
			} else {
				properties.add(property);
			}
		}
		return properties;
	}

	/**
	 * 获取字段
	 */
	private List<Field> getFields(Class<?> type) {
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
