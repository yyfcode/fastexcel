package com.jeeapp.excel.builder;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;

/**
 * @author Justice
 * @since 0.0.2
 */
public class CellBuilder<P extends RowBuilderHelper<P>> extends DataValidationBuilder<CellBuilder<P>, P> {

	private final P parent;

	private RichTextString string;

	private String author;

	private final CellAddress cellAddress;

	private int width = 2;

	private int height = 1;

	protected CellBuilder(P parent, CellAddress cellAddress) {
		super(parent, cellAddress.getRow(), cellAddress.getRow(), cellAddress.getColumn(), cellAddress.getColumn());
		this.parent = parent;
		this.cellAddress = cellAddress;
	}

	public CellBuilder<P> setCellComment(RichTextString string) {
		this.string = string;
		return this;
	}

	public CellBuilder<P> setCommentText(String text) {
		this.string = creationHelper.createRichTextString(text);
		return this;
	}

	public CellBuilder<P> setCommentAuthor(String author) {
		this.author = author;
		return this;
	}

	public CellBuilder<P> setCommentSize(int width, int height) {
		this.width = width;
		this.height = height;
		return this;
	}

	public CellBuilder<P> addPicture(byte[] pictureData, int format) {
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(cellAddress.getRow());
		clientAnchor.setCol1(cellAddress.getColumn());
		clientAnchor.setRow2(cellAddress.getRow() + 1);
		clientAnchor.setCol2(cellAddress.getColumn() + 1);
		int pictureIndex = parent.workbook.addPicture(pictureData, format);
		parent.sheet.getDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
		return self();
	}

	public CellBuilder<P> setCellValue(Object value) {
		parent.createCell(cellAddress.getRow(), cellAddress.getColumn(), value);
		return this;
	}

	@Override
	public P end() {
		if (string != null && StringUtils.isNotBlank(string.getString())) {
			ClientAnchor clientAnchor = creationHelper.createClientAnchor();
			clientAnchor.setRow1(cellAddress.getRow());
			clientAnchor.setCol1(cellAddress.getColumn());
			clientAnchor.setRow2(cellAddress.getRow() + width);
			clientAnchor.setCol2(cellAddress.getColumn() + height);
			Drawing<?> drawing = parent.sheet.getDrawingPatriarch();
			Comment cellComment = drawing.createCellComment(clientAnchor);
			cellComment.setString(string);
			cellComment.setAuthor(author);
		}
		return super.end();
	}
}
