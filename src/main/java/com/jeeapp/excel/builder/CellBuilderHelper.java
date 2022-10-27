package com.jeeapp.excel.builder;

import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.function.Predicate;

import com.jeeapp.excel.util.CellUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author justice
 */
public abstract class CellBuilderHelper<B extends CellBuilderHelper<B>> {

	private final CommonProperties properties;

	private final Workbook workbook;

	public CellBuilderHelper(Workbook workbook) {
		this.workbook = workbook;
		this.properties = new CommonProperties();
	}

	protected CellBuilderHelper(CellBuilderHelper<?> parent) {
		this.workbook = parent.workbook;
		this.properties = new CommonProperties(parent.properties);
	}

	/**
	 * 匹配满足条件的单元格
	 */
	public CellStyleBuilder<B> matchingCell(Predicate<Cell> predicate) {
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return new CellStyleBuilder<>(result, predicate);
	}

	/**
	 * 匹配列
	 */
	public CellStyleBuilder<B> matchingColumn(int column) {
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return new CellStyleBuilder<>(result, column);
	}

	/**
	 * 匹配所有单元格
	 */
	public CellStyleBuilder<B> matchingAll() {
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return new CellStyleBuilder<>(result);
	}

	/**
	 * 匹配指定区域
	 * @param firstRow 起始行
	 * @param lastRow 结束行必须等于或大于 {@code firstRow}
	 * @param firstCol 起始列
	 * @param lastCol 结束列必须等于或大于 {@code firstCol}
	 * @return
	 */
	public CellStyleBuilder<B> matchingRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return new CellStyleBuilder<>(result, cell -> cell.getColumnIndex() >= firstCol
			&& cell.getColumnIndex() <= lastCol
			&& cell.getRowIndex() >= firstRow
			&& cell.getRowIndex() <= lastRow);
	}

	/**
	 * 行高
	 */
	public B setDefaultRowHeight(int height) {
		properties.height = height;
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return result;
	}

	/**
	 * 列宽
	 */
	public B setDefaultColumnWidth(int width) {
		properties.width = width;
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return result;
	}

	/**
	 * 列宽自适应
	 */
	public B autoSizeColumns(Integer... columns) {
		CollectionUtils.addAll(properties.autoSizeColumns, columns);
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return result;
	}

	/**
	 * 指定列宽
	 */
	public B setColumnWidth(int column, int width) {
		properties.columnWidths.put(column, width);
		@SuppressWarnings("unchecked")
		B result = (B) this;
		return result;
	}

	/**
	 * 设置单元格样式
	 */
	protected void setCellStyle(Cell cell) {
		Map<String, Object> properties = new HashMap<>(this.properties.commonStyles);
		for (Predicate<Cell> predicate : this.properties.cellStyles.keySet()) {
			if (predicate.test(cell)) {
				properties.putAll(this.properties.cellStyles.get(predicate));
			}
		}
		CellUtils.setCellStyleProperties(cell, properties);
	}

	/**
	 * 设置列样式
	 */
	protected void setColumnStyle(Sheet sheet, int column) {
		Map<String, Object> properties = new HashMap<>(this.properties.commonStyles);
		properties.putAll(this.properties.columnStyles.get(column));
		CellUtils.setColumnStyleProperties(sheet, column, properties);
	}

	/**
	 * init sheet
	 */
	protected void initSheet(Sheet sheet) {
		if (!properties.autoSizeColumns.isEmpty() && sheet instanceof SXSSFSheet) {
			((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		}
		for (Integer column : properties.getAutoSizeColumns()) {
			sheet.autoSizeColumn(column);
		}
		if (properties.width != null) {
			sheet.setDefaultColumnWidth(properties.width);
		}
		if (properties.height != null) {
			sheet.setDefaultRowHeightInPoints(properties.height);
		}
		for (int column : properties.columnWidths.keySet()) {
			sheet.setColumnWidth(column, properties.columnWidths.get(column) * 256);
		}
		for (int column : properties.columnStyles.keySet()) {
			CellUtils.setColumnStyleProperties(sheet, column, properties.columnStyles.get(column));
		}
	}

	/**
	 * init row
	 */
	protected void initRow(Row row) {
		if (properties.height != null) {
			row.setHeightInPoints(properties.height);
		}
	}

	/**
	 * 创建格式化实例
	 */
	protected DataFormat createDataFormat() {
		return workbook.createDataFormat();
	}

	/**
	 * 添加单元格样式
	 */
	protected void addCellStyle(Predicate<Cell> predicate, Map<String, Object> properties) {
		if (this.properties.cellStyles.containsKey(predicate)) {
			this.properties.cellStyles.get(predicate).putAll(properties);
		} else {
			this.properties.cellStyles.put(predicate, properties);
		}
	}

	/**
	 * 添加列样式
	 */
	protected void addColumnStyle(int column, Map<String, Object> properties) {
		if (this.properties.columnStyles.containsKey(column)) {
			this.properties.columnStyles.get(column).putAll(properties);
		} else {
			this.properties.columnStyles.put(column, properties);
		}
		this.addCellStyle(cell -> cell.getColumnIndex() == column, properties);
	}

	/**
	 * 添加全局样式
	 */
	protected void addCommonStyle(Map<String, Object> properties) {
		this.properties.commonStyles.putAll(properties);
	}

	@Getter
	@Setter
	protected static class CommonProperties {

		private Integer height;

		private Integer width;

		private Set<Integer> autoSizeColumns = new HashSet<>();

		private Map<Integer, Integer> columnWidths = new HashMap<>();

		private Map<String, Object> commonStyles = new LinkedHashMap<>();

		private Map<Predicate<Cell>, Map<String, Object>> cellStyles = new LinkedHashMap<>();

		private Map<Integer, Map<String, Object>> columnStyles = new LinkedHashMap<>();

		public CommonProperties() {
		}

		public CommonProperties(CommonProperties properties) {
			this.height = properties.height;
			this.width = properties.width;
			this.autoSizeColumns = new HashSet<>(properties.autoSizeColumns);
			this.columnWidths = new HashMap<>(properties.columnWidths);
			this.commonStyles.putAll(properties.commonStyles);
			for (Predicate<Cell> predicate : properties.cellStyles.keySet()) {
				this.cellStyles.put(predicate, new HashMap<>(properties.cellStyles.get(predicate)));
			}
			for (Integer column : properties.columnStyles.keySet()) {
				this.columnStyles.put(column, new HashMap<>(properties.columnStyles.get(column)));
			}
		}
	}
}
