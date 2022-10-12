package io.github.yyfcode.fastexcel.rowset;

import java.util.Collection;
import java.util.Map;

import org.springframework.util.Assert;
import io.github.yyfcode.fastexcel.model.Comment;
import io.github.yyfcode.fastexcel.model.Row;

/**
 * @author Justice
 */
public class MappingResult<T> implements MappingErrors {

	private static final long serialVersionUID = -3831630617137311129L;

	private final Row row;

	private Map<String, Integer> fieldColumns;

	private T target;

	public MappingResult(Row row) {
		this.row = row;
	}

	public MappingResult(Row row, T target) {
		this.row = row;
		this.target = target;
	}

	public void setFieldColumns(Map<String, Integer> fieldColumns) {
		this.fieldColumns = fieldColumns;
	}

	public T getTarget() {
		return target;
	}

	@Override
	public Row getRow() {
		return row;
	}

	public void addErrors(Collection<Comment> errors) {
		row.addComments(errors);
	}

	@Override
	public void addError(int colNum, String errMsg) {
		row.addComment(colNum, errMsg);
	}

	public void addError(String field, String errMsg) {
		Assert.notNull(fieldColumns, "fieldColumns must not be null");
		row.addComment(fieldColumns.getOrDefault(field, 0), errMsg);
	}

	@Override
	public boolean hasErrors() {
		return row.hasComments();
	}
}
