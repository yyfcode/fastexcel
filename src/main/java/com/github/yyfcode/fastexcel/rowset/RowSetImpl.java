package com.github.yyfcode.fastexcel.rowset;

import lombok.Data;
import com.github.yyfcode.fastexcel.model.Row;

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
