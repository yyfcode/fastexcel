package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Picture;

/**
 * @author Justice
 * @since 0.0.2
 */
public class PictureBuilder<B extends PictureBuilder<B>> extends DataValidationBuilder<B> {

	private final SheetBuilder parent;

	private int pictureIndex = 0;

	private ClientAnchor clientAnchor;

	private Double scaleX;

	private Double scaleY;

	public PictureBuilder(SheetBuilder parent, int firstRow, int lastRow, int firstCol, int lastCol) {
		super(parent, firstRow, lastRow, firstCol, lastCol);
		this.parent = parent;
		this.clientAnchor = parent.workbook.getCreationHelper().createClientAnchor();
		this.clientAnchor.setRow1(firstRow);
		this.clientAnchor.setCol1(firstCol);
		this.clientAnchor.setRow2(lastRow + 1);
		this.clientAnchor.setCol2(lastCol + 1);
	}

	public B addPicture(byte[] pictureData, int format) {
		this.pictureIndex = parent.workbook.addPicture(pictureData, format);
		return self();
	}

	public B resizePicture(double scaleX, double scaleY) {
		this.scaleX = scaleX;
		this.scaleY = scaleY;
		return self();
	}

	public B setPictureAnchorType(AnchorType anchorType) {
		this.clientAnchor.setAnchorType(anchorType);
		return self();
	}

	public B setPictureSize(int width, int height) {
		this.clientAnchor.setRow2(region.getLastRow() + width);
		this.clientAnchor.setCol2(region.getLastColumn() + height);
		return self();
	}

	@Override
	@SuppressWarnings("unchecked")
	protected B self() {
		return (B) this;
	}

	@Override
	public SheetBuilder end() {
		if (pictureIndex > 0) {
			Picture picture = parent.sheet.getDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
			if (scaleX != null && scaleY != null) {
				picture.resize(scaleX, scaleY);
			}
		}
		return super.end();
	}
}
