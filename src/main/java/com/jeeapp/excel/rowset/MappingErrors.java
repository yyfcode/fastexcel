package com.jeeapp.excel.rowset;

import java.io.Serializable;

import com.jeeapp.excel.model.Row;

/**
 * @author Justice
 */
public interface MappingErrors extends Serializable {

	Row getRow();

	void addError(int colNum, String errMsg);

	boolean hasErrors();
}
