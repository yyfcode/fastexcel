package com.github.yyfcode.fastexcel.model;

import java.io.Serializable;
import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

import lombok.Data;
import org.apache.commons.collections4.CollectionUtils;

/**
 * @author Justice
 */
@Data
public class Row implements Serializable {

	private static final long serialVersionUID = -8762508052812384289L;

	private int rowNum;

	private String[] cellValues;

	private Set<Comment> comments;

	public Row(int rowNum, String[] cellValues) {
		this(rowNum, cellValues, null);
	}

	public Row(int rowNum, String[] cellValues, Set<Comment> comments) {
		this.rowNum = rowNum;
		this.cellValues = cellValues;
		this.comments = comments;
	}

	public void addComment(int colNum, String text) {
		if (this.comments == null) {
			this.comments = new HashSet<>();
		}
		this.comments.add(new Comment(rowNum, colNum, text));
	}

	public void addComments(Collection<Comment> comments) {
		if (this.comments == null) {
			this.comments = new HashSet<>();
		}
		this.comments.addAll(comments);
	}

	public boolean hasComments() {
		return CollectionUtils.isNotEmpty(comments);
	}
}
