package com.jeeapp.excel.rowset;

import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.poifs.filesystem.FileMagic;

/**
 * @author Justice
 */
public interface RowSetReader extends Iterable<RowSet> {

	RowSet read() throws Exception;

	static RowSetReader open(InputStream inputStream) throws Exception {
		InputStream in = FileMagic.prepareToCheckMagic(inputStream);
		FileMagic fileMagic = FileMagic.valueOf(in);
		if (fileMagic == FileMagic.OLE2) {
			return new EventXlsRowSetReader(in);
		} else if (fileMagic == FileMagic.OOXML) {
			return new StreamingXlsxRowSetReader(in);
		} else {
			throw new IllegalStateException("Your file appears not to be a valid excel file");
		}
	}

	@Override
	default Iterator<RowSet> iterator() {
		return new Iterator<RowSet>() {

			private RowSet rowSet;

			@Override
			public boolean hasNext() {
				try {
					rowSet = read();
					return rowSet != null;
				} catch (Exception e) {
					return false;
				}
			}

			@Override
			public RowSet next() {
				return rowSet;
			}
		};
	}
}
