package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Justice
 */
public class PictureBuilder<P> {

	private final P parent;

	private final Sheet sheet;

	private final int pictureIndex;

	private ClientAnchor clientAnchor;

	public PictureBuilder(P parent, Sheet sheet, int pictureIndex) {
		this.parent = parent;
		this.sheet = sheet;
		this.pictureIndex = pictureIndex;
		this.clientAnchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
	}

	public PictureBuilder<P> setRow1(int row1) {
		this.clientAnchor.setRow1(row1);
		return this;
	}

	public PictureBuilder<P> setCol1(int col1) {
		this.clientAnchor.setCol1(col1);
		return this;
	}

	public PictureBuilder<P> setRow2(int row2) {
		this.clientAnchor.setRow2(row2);
		return this;
	}

	public PictureBuilder<P> setCol2(int col2) {
		this.clientAnchor.setCol2(col2);
		return this;
	}

	public PictureBuilder<P> setSize(int width, int height) {
		this.clientAnchor.setCol1(clientAnchor.getCol1());
		this.clientAnchor.setRow1(clientAnchor.getRow1());
		this.clientAnchor.setCol2(clientAnchor.getCol1() + height);
		this.clientAnchor.setRow2(clientAnchor.getRow1() + width);
		return this;
	}

	public PictureBuilder<P> setDx1(int dx1) {
		this.clientAnchor.setDx1(dx1);
		return this;
	}

	public PictureBuilder<P> setDx2(int dx2) {
		this.clientAnchor.setDx2(dx2);
		return this;
	}

	public PictureBuilder<P> setDy1(int dy1) {
		this.clientAnchor.setDy1(dy1);
		return this;
	}

	public PictureBuilder<P> setDy2(int dy2) {
		this.clientAnchor.setDy2(dy2);
		return this;
	}

	public PictureBuilder<P> setAnchorType(AnchorType anchorType) {
		this.clientAnchor.setAnchorType(anchorType);
		return this;
	}

	public PictureBuilder<P> setClientAnchor(ClientAnchor clientAnchor) {
		this.clientAnchor = clientAnchor;
		return this;
	}

	public P insert() {
		sheet.getDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
		return parent;
	}
}
