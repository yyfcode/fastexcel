package io.github.yyfcode.fastexcel.rowset;

import java.io.Serializable;

import io.github.yyfcode.fastexcel.model.Row;

/**
 * @author Justice
 */
public interface RowSet extends Serializable {

	int getSheetIndex();

	String getSheetName();

	int getLastRowNum();

	Row getRow();
}
