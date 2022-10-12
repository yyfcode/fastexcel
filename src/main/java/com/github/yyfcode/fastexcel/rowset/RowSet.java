package com.github.yyfcode.fastexcel.rowset;

import java.io.Serializable;

import com.github.yyfcode.fastexcel.model.Row;

/**
 * @author Justice
 */
public interface RowSet extends Serializable {

	int getSheetIndex();

	String getSheetName();

	int getLastRowNum();

	Row getRow();
}
