package com.jeeapp.excel.model;

import java.lang.reflect.Field;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * @author Justice
 */
@Getter
@Setter
@ToString
@EqualsAndHashCode(callSuper = true)
public class Column extends Cell {

	private final static Integer DEFAULT_WIDTH = 8;

	private Integer width = DEFAULT_WIDTH;

	private Boolean hidden;

	public Column(Field field) {
		super(field);
		this.setWidth(property.width());
		this.setHidden(property.hidden());
	}
}
