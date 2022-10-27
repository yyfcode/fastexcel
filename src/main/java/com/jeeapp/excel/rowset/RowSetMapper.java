package com.jeeapp.excel.rowset;

/**
 * @author Justice
 */
public interface RowSetMapper<T> {

	T mapRowSet(RowSet rowSet) throws MappingException;

	default MappingResult<T> getMappingResult(RowSet rowSet) {
		try {
			return new MappingResult<>(rowSet.getRow(), mapRowSet(rowSet));
		} catch (MappingException ex) {
			return new MappingResult<>(ex.getRow());
		}
	}
}
