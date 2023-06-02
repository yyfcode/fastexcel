package com.jeeapp.excel.builder;

/**
 * @author Justice
 */

import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author Justice
 * @since 0.0.4
 */
public final class SheetHelper {

	private final Sheet sheet;

	private final Drawing<?> drawing;

	private final CreationHelper creationHelper;

	private final DataValidationHelper dataValidationHelper;

	private final SpreadsheetVersion spreadsheetVersion;

	SheetHelper(Sheet sheet) {
		this.sheet = sheet;
		this.drawing = sheet.createDrawingPatriarch();
		this.dataValidationHelper = sheet.getDataValidationHelper();
		this.creationHelper = sheet.getWorkbook().getCreationHelper();
		this.spreadsheetVersion = sheet.getWorkbook().getSpreadsheetVersion();
	}

	public DataValidationHelper getDataValidationHelper() {
		return dataValidationHelper;
	}

	public Integer getMaxRows() {
		return spreadsheetVersion.getMaxRows();
	}

	public Integer getMaxColumns() {
		return spreadsheetVersion.getMaxColumns();
	}

	public RichTextString createRichTextString(String text) {
		return creationHelper.createRichTextString(text);
	}

	public Hyperlink createHyperlink(HyperlinkType type) {
		return creationHelper.createHyperlink(type);
	}

	public ClientAnchor createClientAnchor() {
		return creationHelper.createClientAnchor();
	}

	public Integer addPicture(byte[] pictureData, int format) {
		return sheet.getWorkbook().addPicture(pictureData, format);
	}

	public Font createFont(Map<String, Object> properties) {
		return CellUtils.getFont(sheet.getWorkbook(), properties);
	}

	public Comment createCellComment(ClientAnchor clientAnchor) {
		return drawing.createCellComment(clientAnchor);
	}

	public Picture createPicture(ClientAnchor clientAnchor, int pictureIndex) {
		return drawing.createPicture(clientAnchor, pictureIndex);
	}

	public void addValidationData(DataValidation validation) {
		sheet.addValidationData(validation);
	}
}
