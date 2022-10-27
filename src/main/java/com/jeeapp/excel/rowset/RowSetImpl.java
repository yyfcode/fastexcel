package com.jeeapp.excel.rowset;

import lombok.Data;
import com.jeeapp.excel.model.Row;

/**
 * @author Justice
 */
@Data
class RowSetImpl implements RowSet {

	private static final long serialVersionUID = 1126867710687716611L;

	private int sheetIndex;

	private String sheetName;

	private int lastRowNum;

	private Row row;
}
