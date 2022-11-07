package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;

/**
 * @author Justice
 */
public class CellCommentBuilder<P> {

	private final P parent;

	private final Sheet sheet;

	private ClientAnchor clientAnchor;

	private int row;

	private int column;

	private RichTextString string;

	private String author;

	private CellAddress cellAddress;

	public CellCommentBuilder(P parent, Sheet sheet, String text) {
		this.parent = parent;
		this.sheet = sheet;
		CreationHelper creationHelper = sheet.getWorkbook().getCreationHelper();
		this.clientAnchor = creationHelper.createClientAnchor();
		this.string = creationHelper.createRichTextString(text);
	}

	public CellCommentBuilder<P> setRow1(int row1) {
		this.clientAnchor.setRow1(row1);
		return this;
	}

	public CellCommentBuilder<P> setCol1(int col1) {
		this.clientAnchor.setCol1(col1);
		return this;
	}

	public CellCommentBuilder<P> setRow2(int row2) {
		this.clientAnchor.setRow2(row2);
		return this;
	}

	public CellCommentBuilder<P> setCol2(int col2) {
		this.clientAnchor.setCol2(col2);
		return this;
	}

	public CellCommentBuilder<P> setSize(int width, int height) {
		this.clientAnchor.setCol1(clientAnchor.getCol1());
		this.clientAnchor.setRow1(clientAnchor.getRow1());
		this.clientAnchor.setCol2(clientAnchor.getCol1() + height);
		this.clientAnchor.setRow2(clientAnchor.getRow1() + width);
		return this;
	}

	public CellCommentBuilder<P> setDx1(int dx1) {
		this.clientAnchor.setDx1(dx1);
		return this;
	}

	public CellCommentBuilder<P> setDx2(int dx2) {
		this.clientAnchor.setDx2(dx2);
		return this;
	}

	public CellCommentBuilder<P> setDy1(int dy1) {
		this.clientAnchor.setDy1(dy1);
		return this;
	}

	public CellCommentBuilder<P> setDy2(int dy2) {
		this.clientAnchor.setDy2(dy2);
		return this;
	}

	public CellCommentBuilder<P> setAnchorType(AnchorType anchorType) {
		this.clientAnchor.setAnchorType(anchorType);
		return this;
	}

	public CellCommentBuilder<P> setString(RichTextString string) {
		this.string = string;
		return this;
	}

	public CellCommentBuilder<P> setAuthor(String author) {
		this.author = author;
		return this;
	}

	public CellCommentBuilder<P> setAddress(CellAddress cellAddress) {
		this.cellAddress = cellAddress;
		return this;
	}

	public CellCommentBuilder<P> setClientAnchor(ClientAnchor clientAnchor) {
		this.clientAnchor = clientAnchor;
		return this;
	}

	public P insert() {
		Drawing<?> drawing = sheet.getDrawingPatriarch();
		Comment cellComment = drawing.createCellComment(clientAnchor);
		cellComment.setString(string);
		cellComment.setAuthor(author);
		if (cellAddress != null) {
			cellComment.setAddress(cellAddress);
		}
		return parent;
	}
}
