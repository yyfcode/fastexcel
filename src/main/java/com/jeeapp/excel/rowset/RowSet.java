package com.jeeapp.excel.rowset;

import java.io.Serializable;

import com.jeeapp.excel.model.Row;

/**
 * @author Justice
 */
public interface RowSet extends Serializable {

	int getSheetIndex();

	String getSheetName();

	int getLastRowNum();

	Row getRow();
}
