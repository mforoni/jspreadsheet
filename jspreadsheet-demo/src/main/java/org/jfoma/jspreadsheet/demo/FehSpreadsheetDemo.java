package org.jfoma.jspreadsheet.demo;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import org.apache.logging.log4j.util.Strings;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jspreadsheet.Sheet;
import com.github.mforoni.jspreadsheet.Spreadsheet;
import com.github.mforoni.jspreadsheet.Spreadsheets;

/**
 * @author Foroni Marco
 */
final class FehSpreadsheetDemo {

	private static final Logger LOGGER = LoggerFactory.getLogger(FehSpreadsheetDemo.class);
	private static final String INPUT_FILENAME = "RawData.ods";
	private static final String SHEET_NAME = "Sheet1";
	private static final Path OUTPUT_DIR = Paths.get("output", "spreadsheets");
	private static final String OUTPUT_FILENAME = "FehGamepress.ods";
	private static final String STATS = "Stats";
	private static final String[] HEADER = { "Name", "HP", "ATK", "SPD", "DEF", "RES", "Total", "Rating" };

	private FehSpreadsheetDemo() {}

	private static void run() throws IOException {
		Files.createDirectories(OUTPUT_DIR);
		try (final Spreadsheet out = Spreadsheets.create(Paths.get(OUTPUT_DIR.toString(), OUTPUT_FILENAME).toFile());
				final Spreadsheet in = Spreadsheets.open(JFiles.fromResource(INPUT_FILENAME))) {
			final Sheet sheetIn = in.getSheet(SHEET_NAME);
			final int rows = sheetIn.getRows();
			final Sheet sheetOut = out.addSheet(STATS);
			LOGGER.info("Reading file {}: rows = {}", INPUT_FILENAME, rows);
			int counter = 0;
			sheetOut.setRow(counter++, HEADER);
			final List<Object> values = new ArrayList<>();
			for (int r = 0; r < rows; r++) {
				final Object[] row = sheetIn.getRow(r);
				// JLogger.info(row);
				if (row[0] instanceof String) {
					final String first = (String) row[0];
					if (Strings.isBlank(first)) {
						// do nothing
						LOGGER.info("Writing {}", values);
						sheetOut.setRow(counter++, values.toArray());
						values.clear();
					} else if (first.equals("HP")) {
						// do nothing
					} else {
						values.add(first);
					}
				} else {
					for (final Object obj : row) {
						if (obj != null) {
							values.add(obj);
						}
					}
				}
			}
			out.write();
			LOGGER.info("File {} has been successfully written into {}.", OUTPUT_FILENAME, OUTPUT_DIR);
		}
	}

	public static void main(final String[] args) {
		try {
			run();
		} catch (final Throwable t) {
			t.printStackTrace();
		}
	}
}