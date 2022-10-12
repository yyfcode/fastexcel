package io.github.yyfcode.fastexcel.rowset;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.atomic.AtomicInteger;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.EOFRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * @author Justice
 */
@Slf4j
class EventXlsRowSetReader implements RowSetReader {

	private final AtomicInteger lastRowNum = new AtomicInteger();

	private BlockingQueue<RowSet> rowSetQueue;

	private ExecutorService taskExecutor;

	private InputStream inputStream;

	private FormatTrackingHSSFListener formatListener;

	private boolean open;

	public EventXlsRowSetReader(InputStream inputStream) {
		this.inputStream = inputStream;
	}

	public void open() throws Exception {
		open = true;
		final POIFSFileSystem fileSystem = new POIFSFileSystem(inputStream);
		final HSSFEventFactory hssfEventFactory = new HSSFEventFactory();
		final HSSFRequest request = new HSSFRequest();
		final MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(new HSSFListenerImpl(this));
		formatListener = new FormatTrackingHSSFListener(listener);
		request.addListenerForAllRecords(new SheetRecordCollectingListener(formatListener));
		rowSetQueue = new LinkedBlockingQueue<>();
		taskExecutor = Executors.newSingleThreadExecutor();
		taskExecutor.execute(() -> {
			try {
				hssfEventFactory.processWorkbookEvents(request, fileSystem);
			} catch (IOException ignored) {
			}
		});
	}

	@Override
	public RowSet read() throws Exception {
		if (!open) {
			open();
		}
		RowSet rowSet = rowSetQueue.take();
		if (rowSet.getRow() == null) {
			close();
			return null;
		}
		return RowSetBuilder.builder(rowSet).withLastRowNum(lastRowNum.get()).build();
	}

	public void close() throws Exception {
		if (inputStream != null) {
			inputStream.close();
			inputStream = null;
		}
		if (taskExecutor != null) {
			taskExecutor.shutdown();
			this.taskExecutor = null;
		}
	}

	/**
	 * @author yinyf
	 */
	static final class HSSFListenerImpl implements HSSFListener {

		private final EventXlsRowSetReader rowSetReader;

		private final RowSetBuilder rowSetBuilder;

		private final List<String> cellValues;

		private List<BoundSheetRecord> boundSheetRecords;

		private BoundSheetRecord[] orderedBoundSheetRecords;

		private SSTRecord sstRecord;

		private boolean worksheet;

		private int eofCount = -1;

		private int sheetIndex = -1;

		private boolean outputNextStringRecord;

		HSSFListenerImpl(EventXlsRowSetReader rowSetReader) {
			this.rowSetReader = rowSetReader;
			this.rowSetBuilder = RowSetBuilder.builder();
			this.cellValues = new ArrayList<>();
		}

		@Override
		public void processRecord(Record record) {
			String thisStr = null;
			switch (record.getSid()) {
				case BoundSheetRecord.sid:
					if (boundSheetRecords == null) {
						boundSheetRecords = new ArrayList<>();
					}
					boundSheetRecords.add((BoundSheetRecord) record);
					break;
				case BOFRecord.sid:
					BOFRecord bofRecord = (BOFRecord) record;
					if (bofRecord.getType() == BOFRecord.TYPE_WORKSHEET) {
						worksheet = true;
						sheetIndex++;
						if (orderedBoundSheetRecords == null) {
							orderedBoundSheetRecords = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
						}
						rowSetBuilder.withSheet(sheetIndex, orderedBoundSheetRecords[sheetIndex].getSheetname());
					} else {
						worksheet = false;
					}
					break;
				case SSTRecord.sid:
					sstRecord = (SSTRecord) record;
					break;
				case BlankRecord.sid:
					thisStr = "";
					break;
				case BoolErrRecord.sid:
					BoolErrRecord boolErrRecord = (BoolErrRecord) record;
					thisStr = boolErrRecord.getBooleanValue() + "";
					break;
				case FormulaRecord.sid:
					FormulaRecord formulaRecord = (FormulaRecord) record;
					if (formulaRecord.hasCachedResultString() && Double.isNaN(formulaRecord.getValue())) {
						outputNextStringRecord = true;
					} else {
						thisStr = rowSetReader.formatListener.formatNumberDateCell(formulaRecord);
					}
					break;
				case StringRecord.sid:
					if (outputNextStringRecord) {
						StringRecord stringRecord = (StringRecord) record;
						thisStr = stringRecord.getString();
						outputNextStringRecord = false;
					}
					break;
				case LabelRecord.sid:
					LabelRecord labelRecord = (LabelRecord) record;
					thisStr = labelRecord.getValue();
					break;
				case LabelSSTRecord.sid:
					LabelSSTRecord labelSstRecord = (LabelSSTRecord) record;
					if (sstRecord == null) {
						thisStr = "";
					} else {
						thisStr = sstRecord.getString(labelSstRecord.getSSTIndex()).toString();
					}
					break;
				case NumberRecord.sid:
					NumberRecord numberRecord = (NumberRecord) record;
					thisStr = rowSetReader.formatListener.formatNumberDateCell(numberRecord);
					break;
				case EOFRecord.sid:
					eofCount++;
					if (worksheet && eofCount == boundSheetRecords.size()) {
						try {
							rowSetReader.rowSetQueue.put(rowSetBuilder.withNullRow().build());
						} catch (Exception e) {
							log.error("Unable to send row to the queue", e);
						}
					}
					break;
				default:
					break;
			}
			if (thisStr != null) {
				cellValues.add(thisStr);
			}
			if (record instanceof LastCellOfRowDummyRecord) {
				LastCellOfRowDummyRecord lastCellOfRowDummyRecord = (LastCellOfRowDummyRecord) record;
				int rowNum = lastCellOfRowDummyRecord.getRow();
				try {
					rowSetReader.lastRowNum.incrementAndGet();
					String[] cellValues = this.cellValues.toArray(new String[]{});
					rowSetReader.rowSetQueue.put(rowSetBuilder.withRow(rowNum, cellValues).build());
				} catch (Exception e) {
					log.error("Unable to send row to the queue", e);
				}
				cellValues.clear();
			}
		}
	}
}
