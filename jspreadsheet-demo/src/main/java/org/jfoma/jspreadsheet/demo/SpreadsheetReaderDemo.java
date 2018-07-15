package org.jfoma.jspreadsheet.demo;

import java.io.File;
import java.io.IOException;
import javax.annotation.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jbasic.util.JLogger;
import com.github.mforoni.jbasic.util.JTest;
import com.github.mforoni.jspreadsheet.Sheet;
import com.github.mforoni.jspreadsheet.Spreadsheet;
import com.github.mforoni.jspreadsheet.Spreadsheets;
import com.github.mforoni.jspreadsheet.Spreadsheets.API;

final class SpreadsheetReaderDemo {

	private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetReaderDemo.class);
	private static final String EMPTY_XLS = "Empty.xls";
	private static final String RATINGS_XLS_SAMPLE = "Ratings.xls";
	// XLSX
	private static final String FINANCIAL_XLSX = "Financial.xlsx";
	private static final String SALES_XLSX = "Sales.xlsx";
	private static final String RATINGS_XLSX_SAMPLE = "Ratings.xlsx";
	// ODS
	private static final String STOCK_CHART_ODS = "StockChart.ods";
	private static final String FEHTEST_ODS = "FehTest.ods";
	private static final String RATINGS_ODS_SAMPLE = "Ratings.ods";

	private SpreadsheetReaderDemo() {}

	static void readSpreadsheet(final String filename) throws IOException {
		readSpreadsheet(filename, null);
	}

	static void readSpreadsheet(final String filename, @Nullable final API api) throws IOException {
		final File file = JFiles.fromResource(filename);
		try (final Spreadsheet spreadsheet = api == null ? Spreadsheets.open(file) : Spreadsheets.open(file, api)) {
			for (final String sheetName : spreadsheet.getSheetNames()) {
				final Sheet sheet = spreadsheet.getSheet(sheetName);
				final int rows = sheet.getRows();
				final int cols = sheet.getColumns();
				LOGGER.info("Reading file {} sheet {}: rows = {}, columns = {}", filename, sheetName, rows, cols);
				for (int r = 0; r < rows; r++) {
					try {
						final Object[] row = sheet.getRow(r);
						JLogger.info(row);
					} catch (final Throwable t) {
						LOGGER.error("Error while reading row {} of file {}", r, filename, t);
					}
				}
			}
		}
	}

	private static JTest readingTest(final String filename, final API api) {
		return new JTest("Test reading file ".concat(filename).concat(" using API ").concat(api.name())) {

			@Override
			protected void run() {
				try {
					readSpreadsheet(filename, api);
				} catch (final IOException e) {
					LOGGER.error("Error:", e);
				}
			}
		};
	}

	private static JTest readingTest(final String filename) {
		return new JTest("Test reading file ".concat(filename)) {

			@Override
			protected void run() {
				try {
					readSpreadsheet(filename);
				} catch (final IOException e) {
					LOGGER.error("Error:", e);
				}
			}
		};
	}

	private static void readExcels() {
		final JTest test1 = readingTest(FINANCIAL_XLSX).start();
		final JTest test2 = readingTest(SALES_XLSX).start();
		final JTest test3 = readingTest(EMPTY_XLS).start();
		final JTest test4 = readingTest(EMPTY_XLS, API.JXL).start();
		final JTest test5 = readingTest(RATINGS_XLS_SAMPLE).start();
		final JTest test6 = readingTest(RATINGS_XLS_SAMPLE, API.JXL).start();
		test1.log();
		test2.log();
		test3.log();
		test4.log();
		test5.log();
		test6.log();
	}

	private static void readOds() {
		final JTest test1 = readingTest(STOCK_CHART_ODS).start();
		final JTest test2 = readingTest(FEHTEST_ODS).start();
		test1.log();
		test2.log();
	}

	private static void readXlsxAndOds() {
		final JTest test1 = readingTest(RATINGS_XLSX_SAMPLE).start();
		final JTest test2 = readingTest(RATINGS_ODS_SAMPLE).start();
		test1.log();
		test2.log();
	}

	public static void main(final String[] args) {
		try {
			readExcels();
			readOds();
			readXlsxAndOds();
		} catch (final Throwable t) {
			t.printStackTrace();
		}
	}
}