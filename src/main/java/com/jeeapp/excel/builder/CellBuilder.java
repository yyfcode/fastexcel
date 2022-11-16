package com.jeeapp.excel.builder;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;

/**
 * @author Justice
 * @since 0.0.2
 */
public class CellBuilder extends PictureBuilder<CellBuilder> {

	private final SheetBuilder parent;

	private final CreationHelper creationHelper;

	private RichTextString string;

	private String author;

	private ClientAnchor clientAnchor;

	public CellBuilder(SheetBuilder parent, CellAddress cellAddress) {
		super(parent, cellAddress.getRow(), cellAddress.getRow(), cellAddress.getColumn(), cellAddress.getColumn());
		this.parent = parent;
		this.creationHelper = parent.sheet.getWorkbook().getCreationHelper();
		this.clientAnchor = creationHelper.createClientAnchor();
		this.clientAnchor.setRow1(cellAddress.getRow());
		this.clientAnchor.setCol1(cellAddress.getColumn());
	}

	public CellBuilder setCellComment(RichTextString string) {
		this.string = string;
		return this;
	}

	public CellBuilder setCommentText(String text) {
		this.string = creationHelper.createRichTextString(text);
		return this;
	}

	public CellBuilder setCommentAuthor(String author) {
		this.author = author;
		return this;
	}

	public CellBuilder setCommentAnchorType(AnchorType anchorType) {
		this.clientAnchor.setAnchorType(anchorType);
		return this;
	}

	public CellBuilder setCommentSize(int width, int height) {
		this.clientAnchor.setRow2(region.getLastRow() + width);
		this.clientAnchor.setCol2(region.getLastColumn() + height);
		return this;
	}

	public SheetBuilder setCellValue(Object value) {
		return end().createCell(region.getFirstRow(), region.getFirstColumn(), value);
	}

	@Override
	public SheetBuilder end() {
		if (string != null && StringUtils.isNotBlank(string.getString())) {
			Drawing<?> drawing = parent.sheet.getDrawingPatriarch();
			Comment cellComment = drawing.createCellComment(clientAnchor);
			cellComment.setString(string);
			cellComment.setAuthor(author);
		}
		return super.end();
	}
}
