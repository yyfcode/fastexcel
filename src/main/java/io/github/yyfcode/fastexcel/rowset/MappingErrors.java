package io.github.yyfcode.fastexcel.rowset;

import java.io.Serializable;

import io.github.yyfcode.fastexcel.model.Row;

/**
 * @author Justice
 */
public interface MappingErrors extends Serializable {

	Row getRow();

	void addError(int colNum, String errMsg);

	boolean hasErrors();
}
