package io.github.yyfcode.fastexcel.model;

import java.io.Serializable;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * @author Justice
 */
@Getter
@Setter
@ToString
@EqualsAndHashCode(of = {"rowNum", "colNum"}, callSuper = false)
public class Comment implements Serializable {

	private static final long serialVersionUID = 2778225839846005500L;

	private int rowNum;

	private int colNum;

	private String text;

	private String author;

	public Comment(int rowNum, int colNum, String text) {
		this(rowNum, colNum, text, "");
	}

	public Comment(int rowNum, int colNum, String text, String author) {
		this.rowNum = rowNum;
		this.colNum = colNum;
		this.text = text;
		this.author = author;
	}
}
