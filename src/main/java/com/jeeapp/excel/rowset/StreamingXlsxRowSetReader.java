package com.jeeapp.excel.rowset;

import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import lombok.extern.slf4j.Slf4j;
import org.xml.sax.Attributes;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.springframework.util.xml.StaxUtils;

/**
 * @author Justice
 */
@Slf4j
class StreamingXlsxRowSetReader implements RowSetReader {

	private final RowSetBuilder rowSetBuilder;

	private XSSFReader.SheetIterator sheetIterator;

	private SharedStrings sharedStrings;

	private Styles styles;

	private InputStream inputStream;

	private XMLStreamReader xmlStreamReader;

	private ValueRetrievingContentsHandler contentHandler;

	private XSSFSheetXMLHandler sheetHandler;

	private boolean open;

	private int sheetIndex;

	public StreamingXlsxRowSetReader(InputStream inputStream) {
		this.inputStream = inputStream;
		this.open = false;
		this.rowSetBuilder = RowSetBuilder.builder();
	}

	protected boolean nextSheet() throws Exception {
		if (sheetIterator.hasNext()) {
			InputStream is = sheetIterator.next();
			String sheetName = sheetIterator.getSheetName();
			contentHandler = new ValueRetrievingContentsHandler();
			sheetHandler = new XSSFSheetXMLHandler(styles, sharedStrings, contentHandler, false);
			xmlStreamReader = StaxUtils.createDefensiveInputFactory().createXMLStreamReader(is);
			sheetIndex++;
			rowSetBuilder.withSheet(sheetIndex, sheetName);
			return true;
		}
		return false;
	}

	public void open() throws Exception {
		open = true;
		ZipSecureFile.setMinInflateRatio(0);
		OPCPackage opcPackage = OPCPackage.open(inputStream);
		XSSFReader reader = new XSSFReader(opcPackage);
		sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
		sharedStrings = new ReadOnlySharedStringsTable(opcPackage);
		styles = reader.getStylesTable();
		sheetIndex = -1;
		nextSheet();
	}

	@Override
	public RowSet read() throws Exception {
		if (!open) {
			open();
		}
		while (xmlStreamReader.hasNext()) {
			int type = xmlStreamReader.next();
			if (type == XMLStreamConstants.START_DOCUMENT) {
				sheetHandler.startDocument();
			} else if (type == XMLStreamConstants.END_DOCUMENT) {
				sheetHandler.endDocument();
				if (nextSheet()) {
					return read();
				}
				close();
				return null;
			} else if (type == XMLStreamConstants.CHARACTERS) {
				int textLength = xmlStreamReader.getTextLength();
				int textStart = xmlStreamReader.getTextStart();
				sheetHandler.characters(xmlStreamReader.getTextCharacters(), textStart, textLength);
			} else if (type == XMLStreamConstants.START_ELEMENT) {
				String localName = xmlStreamReader.getLocalName();
				if ("dimension".equals(localName)) {
					// ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
					String v = xmlStreamReader.getAttributeValue(null, "ref");
					if (v != null && v.indexOf(':') > -1) {
						CellRangeAddress range = CellRangeAddress.valueOf(v);
						int lastRow = range.getLastRow();
						int firstRow = range.getFirstRow();
						rowSetBuilder.withLastRowNum(lastRow - firstRow + 1);
					}
				} else {
					Attributes delegating = new AttributesAdapter(xmlStreamReader);
					sheetHandler.startElement(null, localName, null, delegating);
				}
			} else if (type == XMLStreamConstants.END_ELEMENT) {
				String tag = xmlStreamReader.getLocalName();
				sheetHandler.endElement(null, tag, null);
				if ("row".equals(tag)) {
					return rowSetBuilder.withRow(contentHandler.getRowNum(), contentHandler.getCellValues()).build();
				}
			}
		}
		return null;
	}

	public void close() throws Exception {
		if (xmlStreamReader != null) {
			try {
				xmlStreamReader.close();
			} catch (Exception ignore) {
			}
			xmlStreamReader = null;
		}
		if (inputStream != null) {
			inputStream.close();
			inputStream = null;
		}
	}


	static class ValueRetrievingContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

		private int rowNum;

		private String[] cellValues;

		@Override
		public void startRow(int rowNum) {
			// Prepare for this row
			if (cellValues == null) {
				cellValues = new String[0];
			}
			Arrays.fill(cellValues, null);
			this.rowNum = rowNum;
		}

		@Override
		public void endRow(int rowNum) {
			this.rowNum = rowNum;
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {
			int col = new CellReference(cellReference).getCol();
			if (cellValues.length <= col) {
				String[] newValues = Arrays.copyOf(cellValues, col + 1);
				Arrays.setAll(newValues, (idx) -> newValues[idx]);
				cellValues = newValues;
			}
			cellValues[col] = formattedValue;
		}

		String[] getCellValues() {
			return Arrays.copyOf(cellValues, cellValues.length);
		}

		int getRowNum() {
			return rowNum;
		}
	}

	static final class AttributesAdapter implements Attributes {

		private final Map<String, String> attributes = new HashMap<>();

		private AttributesAdapter(XMLStreamReader delegate) {
			for (int i = 0; i < delegate.getAttributeCount(); i++) {
				String name = delegate.getAttributeLocalName(i);
				String value = delegate.getAttributeValue(i);
				attributes.put(name, value);
			}
		}

		@Override
		public int getLength() {
			return attributes.size();
		}

		@Override
		public String getURI(int index) {
			return null;
		}

		@Override
		public String getLocalName(int index) {
			return null;
		}

		@Override
		public String getQName(int index) {
			return null;
		}

		@Override
		public String getType(int index) {
			return null;
		}

		@Override
		public String getValue(int index) {
			return null;
		}

		@Override
		public int getIndex(String uri, String localName) {
			return 0;
		}

		@Override
		public int getIndex(String qName) {
			return 0;
		}

		@Override
		public String getType(String uri, String localName) {
			return null;
		}

		@Override
		public String getType(String qName) {
			return null;
		}

		@Override
		public String getValue(String uri, String localName) {
			return attributes.get(localName);
		}

		@Override
		public String getValue(String qName) {
			return attributes.get(qName);
		}
	}
}
