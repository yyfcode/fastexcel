package com.jeeapp.excel.builder;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.function.Predicate;

import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author justice
 */
abstract class CellBuilderHelper<B extends CellBuilderHelper<B>> {

	private final CommonProperties properties;

	protected final Workbook workbook;

	protected CellBuilderHelper(Workbook workbook) {
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
	public CellStyleBuilder<?, B> matchingCell(Predicate<Cell> predicate) {
		return new CellStyleBuilder<>(self(), predicate);
	}

	/**
	 * 匹配列
	 */
	public CellStyleBuilder<?, B> matchingColumn(int column) {
		return new CellStyleBuilder<>(self(), (short) column);
	}

	/**
	 * 匹配行
	 */
	public CellStyleBuilder<?, B> matchingRow(int row) {
		return new CellStyleBuilder<>(self(), row);
	}

	/**
	 * 匹配所有单元格
	 */
	public CellStyleBuilder<?, B> matchingAll() {
		return new CellStyleBuilder<>(self());
	}

	/**
	 * 匹配指定区域
	 * @param firstRow 起始行
	 * @param lastRow 结束行必须等于或大于 {@code firstRow}
	 * @param firstCol 起始列
	 * @param lastCol 结束列必须等于或大于 {@code firstCol}
	 */
	public CellStyleBuilder<?, B> matchingRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
		return new CellStyleBuilder<>(self(), firstRow, lastRow, firstCol, lastCol);
	}

	/**
	 * 行高
	 */
	protected B setDefaultRowHeight(int height) {
		properties.height = height;
		return self();
	}

	/**
	 * 列宽
	 */
	protected B setDefaultColumnWidth(int width) {
		properties.width = width;
		return self();
	}

	protected abstract B self();


	/**
	 * 设置表样式
	 */
	protected void setSheetStyle(Sheet sheet) {
		for (int column : properties.columnStyles.keySet()) {
			setColumnStyle(sheet, column);
		}
		for (int rowNum : properties.rowStyles.keySet()) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				row = sheet.createRow(rowNum);
			}
			setRowStyle(sheet, row);
		}
		for (CellRangeAddress region : properties.mergedRegions.getCellRangeAddresses()) {
			sheet.addMergedRegion(region);
			setRegionStyle(sheet, region);
		}
	}

	/**
	 * 设置行样式，行样式并不包含全局样式，需要单独设置
	 */
	private void setRowStyle(Sheet sheet, Row row) {
		int rowNum = row.getRowNum();
		Map<String, Object> properties = new HashMap<>();
		if (this.properties.rowStyles.containsKey(rowNum)) {
			properties.putAll(this.properties.rowStyles.get(rowNum));
		}
		if (MapUtils.isNotEmpty(properties)) {
			CellUtils.setRowStyleProperties(sheet, row, properties);
		}
	}

	/**
	 * 设置列样式，列样式并不包含全局样式，需要单独设置
	 */
	private void setColumnStyle(Sheet sheet, int column) {
		Map<String, Object> properties = new HashMap<>();
		if (this.properties.columnStyles.containsKey(column)) {
			properties.putAll(this.properties.columnStyles.get(column));
		}
		if (MapUtils.isNotEmpty(properties)) {
			CellUtils.setColumnStyleProperties(sheet, column, properties);
		}
	}

	/**
	 * 设置区域样式
	 */
	private void setRegionStyle(Sheet sheet, CellRangeAddress region) {
		Map<String, Object> properties = new HashMap<>(this.properties.commonStyles);
		if (this.properties.regionStyles.containsKey(region)) {
			properties.putAll(this.properties.regionStyles.get(region));
		}
		if (MapUtils.isNotEmpty(properties)) {
			CellUtils.setRegionStyleProperties(sheet, region, properties);
		}
	}

	/**
	 * 设置单元格样式
	 */
	protected void setCellStyle(Cell cell) {
		Map<String, Object> properties = new HashMap<>(this.properties.commonStyles);
		CellAddress cellAddress = new CellAddress(cell);
		if (this.properties.cellStyles.containsKey(cellAddress)) {
			properties.putAll(this.properties.cellStyles.get(cellAddress));
			this.properties.cellStyles.remove(cellAddress);
		}
		for (Predicate<Cell> predicate : this.properties.customStyles.keySet()) {
			if (predicate.test(cell)) {
				properties.putAll(this.properties.customStyles.get(predicate));
			}
		}
		if (MapUtils.isNotEmpty(properties)) {
			CellUtils.setCellStyleProperties(cell, properties);
		}
	}

	/**
	 * init sheet
	 */
	protected void initSheet(Sheet sheet) {
		if (properties.width != null) {
			sheet.setDefaultColumnWidth(properties.width);
		}
		if (properties.height != null) {
			sheet.setDefaultRowHeightInPoints(properties.height);
		}
	}

	/**
	 * 添加自定义样式
	 */
	protected void addCustomStyle(Predicate<Cell> predicate, Map<String, Object> properties) {
		if (this.properties.customStyles.containsKey(predicate)) {
			this.properties.customStyles.get(predicate).putAll(properties);
		} else {
			this.properties.customStyles.put(predicate, properties);
		}
	}

	/**
	 * 添加单元格样式
	 */
	protected void addCellStyle(CellAddress cellAddress, Map<String, Object> properties) {
		if (this.properties.cellStyles.containsKey(cellAddress)) {
			this.properties.cellStyles.get(cellAddress).putAll(properties);
		} else {
			this.properties.cellStyles.put(cellAddress, properties);
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
		this.addCustomStyle(cell -> cell.getColumnIndex() == column, properties);
	}

	/**
	 * 添加行样式
	 */
	protected void addRowStyle(int row, Map<String, Object> properties) {
		if (this.properties.rowStyles.containsKey(row)) {
			this.properties.rowStyles.get(row).putAll(properties);
		} else {
			this.properties.rowStyles.put(row, properties);
		}
		this.addCustomStyle(cell -> cell.getRowIndex() == row, properties);
	}

	/**
	 * 添加全局样式
	 */
	protected void addCommonStyle(Map<String, Object> properties) {
		this.properties.commonStyles.putAll(properties);
	}

	/**
	 * 添加区域样式
	 */
	protected void addRegionStyle(CellRangeAddress region, Map<String, Object> properties) {
		if (this.properties.regionStyles.containsKey(region)) {
			this.properties.regionStyles.get(region).putAll(properties);
		} else {
			this.properties.regionStyles.put(region, properties);
		}
		this.addCustomStyle(cell -> cell.getColumnIndex() >= region.getFirstColumn()
			&& cell.getColumnIndex() <= region.getLastColumn()
			&& cell.getRowIndex() >= region.getFirstRow()
			&& cell.getRowIndex() <= region.getLastRow(), properties);
	}

	/**
	 * 添加全局样式
	 */
	protected void addMergedRegion(CellRangeAddress region) {
		this.properties.mergedRegions.addCellRangeAddress(region);
	}

	@Getter
	@Setter
	protected static class CommonProperties {

		private Integer width;

		private Integer height;

		private Map<String, Object> commonStyles = new LinkedHashMap<>();

		private CellRangeAddressList mergedRegions = new CellRangeAddressList();

		private Map<Predicate<Cell>, Map<String, Object>> customStyles = new LinkedHashMap<>();

		private Map<CellAddress, Map<String, Object>> cellStyles = new LinkedHashMap<>();

		private Map<Integer, Map<String, Object>> columnStyles = new LinkedHashMap<>();

		private Map<Integer, Map<String, Object>> rowStyles = new LinkedHashMap<>();

		private Map<CellRangeAddress, Map<String, Object>> regionStyles = new LinkedHashMap<>();

		public CommonProperties() {
		}

		public CommonProperties(CommonProperties properties) {
			this.width = properties.width;
			this.height = properties.height;
			this.commonStyles.putAll(properties.commonStyles);
			for (Predicate<Cell> predicate : properties.customStyles.keySet()) {
				this.customStyles.put(predicate, new HashMap<>(properties.customStyles.get(predicate)));
			}
			for (CellAddress cellAddress : properties.cellStyles.keySet()) {
				this.cellStyles.put(cellAddress, new HashMap<>(properties.cellStyles.get(cellAddress)));
			}
			for (Integer column : properties.columnStyles.keySet()) {
				this.columnStyles.put(column, new HashMap<>(properties.columnStyles.get(column)));
			}
			for (Integer row : properties.rowStyles.keySet()) {
				this.rowStyles.put(row, new HashMap<>(properties.rowStyles.get(row)));
			}
			for (CellRangeAddress region : properties.regionStyles.keySet()) {
				this.regionStyles.put(region, new HashMap<>(properties.regionStyles.get(region)));
			}
			for (CellRangeAddress region : properties.mergedRegions.getCellRangeAddresses()) {
				this.mergedRegions.addCellRangeAddress(region);
			}
		}
	}
}
