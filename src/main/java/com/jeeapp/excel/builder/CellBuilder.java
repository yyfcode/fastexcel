package com.jeeapp.excel.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.SheetUtil;

/**
 * @author Justice
 * @since 0.0.2
 */
public class CellBuilder<P extends SheetBuilderHelper<P>> extends DataValidationConstraintBuilder<CellBuilder<P>, P> {

	private final P parent;

	private final CellAddress cellAddress;

	protected CellBuilder(P parent, CellAddress cellAddress) {
		super(parent, cellAddress.getRow(), cellAddress.getRow(), cellAddress.getColumn(), cellAddress.getColumn());
		this.parent = parent;
		this.cellAddress = cellAddress;
	}

	public CellBuilder<P> createCellComment(String text) {
		return createCellComment(text, null, 2, 1);
	}

	public CellBuilder<P> createCellComment(String text, String author) {
		return createCellComment(text, author, 2, 1);
	}

	public CellBuilder<P> createCellComment(String text, String author, int width, int height) {
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		RichTextString string = creationHelper.createRichTextString(text);
		clientAnchor.setRow1(cellAddress.getRow());
		clientAnchor.setCol1(cellAddress.getColumn());
		clientAnchor.setRow2(cellAddress.getRow() + width);
		clientAnchor.setCol2(cellAddress.getColumn() + height);
		Drawing<?> drawing = parent.sheet.getDrawingPatriarch();
		Comment cellComment = drawing.createCellComment(clientAnchor);
		cellComment.setString(string);
		cellComment.setAuthor(author);
		return this;
	}

	public CellBuilder<P> createPicture(byte[] pictureData, int format) {
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(cellAddress.getRow());
		clientAnchor.setCol1(cellAddress.getColumn());
		clientAnchor.setRow2(cellAddress.getRow() + 1);
		clientAnchor.setCol2(cellAddress.getColumn() + 1);
		int pictureIndex = parent.workbook.addPicture(pictureData, format);
		parent.sheet.getDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
		return self();
	}

	/**
	 * 设置单元格值
	 */
	public P setCellValue(Object value) {
		return end().createCell(cellAddress, value);
	}

	/**
	 * 设置空单元
	 */
	public P setBlank() {
		return end().createCell(cellAddress);
	}

	/**
	 * 添加样式
	 */
	public P addCellStyle() {
		P parent = end();
		Cell cell = SheetUtil.getCell(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
		if (cell != null) {
			parent.setCellStyle(cell);
		}
		return parent;
	}
}
