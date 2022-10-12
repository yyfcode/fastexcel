package io.github.yyfcode.fastexcel.model;

import java.io.Serializable;

import lombok.Data;

/**
 * @author Justice
 */
@Data
public class Cell implements Serializable {

	private static final long serialVersionUID = 8481559629638325751L;

	/** 单元格对应的字段属性 */
	private String name;

	/** 单元格对应的字段值 */
	private Object value;

	/** 单元格起始行 */
	private int firstRow;

	/** 单元格结束行 */
	private int lastRow;

	/** 单元格起始列 */
	private int firstCol;

	/** 单元格结束列 */
	private int lastCol;

	protected Cell(String name, Object value) {
		this.name = name;
		this.value = value;
	}

	public static Cell of(String name, Object value) {
		return new Cell(name, value);
	}
}
