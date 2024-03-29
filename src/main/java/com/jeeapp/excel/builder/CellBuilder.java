package com.jeeapp.excel.builder;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * @author Justice
 * @since 0.0.2
 */
public class CellBuilder extends CreationBuilder<CellBuilder> {

	private final CellAddress cellAddress;

	protected CellBuilder(SheetBuilder parent, CellAddress cellAddress) {
		super(parent, cellAddress.getRow(), cellAddress.getRow(), cellAddress.getColumn(), cellAddress.getColumn());
		this.cellAddress = cellAddress;
	}

	public CellBuilder createCellComment(String text) {
		return createCellComment(text, null, 2, 1);
	}

	public CellBuilder createCellComment(String text, String author) {
		return createCellComment(text, author, 2, 1);
	}

	public CellBuilder createCellComment(String text, String author, int width, int height) {
		ClientAnchor clientAnchor = parent.helper.createClientAnchor();
		RichTextString string = parent.helper.createRichTextString(text);
		clientAnchor.setRow1(cellAddress.getRow());
		clientAnchor.setCol1(cellAddress.getColumn());
		clientAnchor.setRow2(cellAddress.getRow() + width);
		clientAnchor.setCol2(cellAddress.getColumn() + height);
		Comment cellComment = parent.helper.createCellComment(clientAnchor);
		cellComment.setString(string);
		cellComment.setAuthor(author);
		return this;
	}

	/**
	 * 设置超链接
	 */
	public SheetBuilder createHyperlink(HyperlinkType hyperlinkType, String address) {
		return createHyperlink(hyperlinkType, address, address);
	}

	/**
	 * 设置超链接
	 */
	public SheetBuilder createHyperlink(HyperlinkType hyperlinkType, String address, String label) {
		Hyperlink hyperlink = parent.helper.createHyperlink(hyperlinkType);
		hyperlink.setAddress(address);
		hyperlink.setLabel(label);
		return parent.matchingCell(cellAddress.getRow(), cellAddress.getColumn())
			.setFontColor(IndexedColors.BLUE.index)
			.setUnderline(XSSFFont.U_SINGLE)
			.setCellValue(hyperlink);
	}

	/**
	 * 设置活动单元格的位置
	 */
	public CellBuilder setActiveCell() {
		parent.sheet.setActiveCell(cellAddress);
		return self();
	}

	/**
	 * 给单元格设置值
	 */
	public SheetBuilder setCellValue(Object value) {
		return super.addCellStyle().createCell(cellAddress, value);
	}

	/**
	 * 从单元格中删除公式和值，并将其类型设置为CellType.BLANK 。保留批注和超链接。
	 * 虽然setCellType(CellType)存在，但它是setCellType(CellType.BLANK)的别名。
	 */
	public SheetBuilder setBlank() {
		return super.addCellStyle().createCell(cellAddress);
	}

	/**
	 * 给单元格设置样式
	 */
	public SheetBuilder setCellStyle() {
		SheetBuilder parent = super.addCellStyle();
		Cell cell = SheetUtil.getCellWithMerges(parent.sheet, cellAddress.getRow(), cellAddress.getColumn());
		if (cell != null) {
			parent.setCellStyle(cell);
		}
		return parent;
	}
}
