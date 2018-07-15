package org.jfoma.jspreadsheet.demo;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.github.mforoni.jbasic.MoreInts;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jbasic.util.JDesktop;
import com.github.mforoni.jspreadsheet.SSCellFormat;
import com.github.mforoni.jspreadsheet.SSColor;
import com.github.mforoni.jspreadsheet.SSFont;
import com.github.mforoni.jspreadsheet.Sheet;
import com.github.mforoni.jspreadsheet.Spreadsheet;
import com.github.mforoni.jspreadsheet.Spreadsheets;
import com.github.mforoni.jspreadsheet.Spreadsheets.API;

/**
 * @author Foroni
 */
final class SpreadsheetWriterDemo {

	private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriterDemo.class);
	private static final String NEW_HEROES_POI = "new_heroes_poi.xlsx";
	private static final String HEROES_POI = "heroes_poi.xlsx";
	private static final String NEW_HEROES_JXL = "new_heroes_jxl.xls";
	private static final String HEROES_JXL = "heroes_jxl.xls";
	private static final String NEW_HEROES_ODS = "new_heroes.ods";
	private static final String HEROES_ODS = "heroes.ods";

	private SpreadsheetWriterDemo() {}

	private static void testJXL(final String outputPath) throws IOException {
		create(outputPath, HEROES_JXL, API.JXL, new TestHeroes());
		createFromTemplate(outputPath, HEROES_JXL, NEW_HEROES_JXL, API.JXL, new TestAppendRow());
		edit(outputPath, NEW_HEROES_JXL, API.JXL, new TestNewSheet());
		openTryingEdit(outputPath, API.JXL, NEW_HEROES_JXL);
	}

	private static void testPOI(final String outputPath) throws IOException {
		create(outputPath, HEROES_POI, API.POI, new TestHeroes());
		createFromTemplate(outputPath, HEROES_POI, NEW_HEROES_POI, API.POI, new TestAppendRow());
		edit(outputPath, NEW_HEROES_POI, API.POI, new TestNewSheet());
		openTryingEdit(outputPath, API.POI, NEW_HEROES_POI);
	}

	private static void testODS(final String outputPath) throws IOException {
		create(outputPath, HEROES_ODS, API.SIMPLE_ODF, new TestHeroes());
		createFromTemplate(outputPath, HEROES_ODS, NEW_HEROES_ODS, API.SIMPLE_ODF, new TestAppendRow());
		edit(outputPath, NEW_HEROES_ODS, API.SIMPLE_ODF, new TestNewSheet());
		openTryingEdit(outputPath, API.SIMPLE_ODF, NEW_HEROES_ODS);
	}

	public static void main(final String[] args) throws IOException {
		final String outputPath = "output/spreadsheets";
		deleteAll(outputPath);
		testJXL(outputPath);
		testPOI(outputPath);
		testODS(outputPath);
		JDesktop.open(Paths.get(outputPath, NEW_HEROES_POI).toFile());
		JDesktop.open(Paths.get(outputPath, NEW_HEROES_ODS).toFile());
	}

	private static void deleteAll(final String outputDir) throws IOException {
		final Path dirPath = Paths.get(outputDir);
		if (Files.exists(dirPath)) {
			List<Path> paths = JFiles.list(dirPath, Spreadsheets.IS_SPREADSHEET);
			LOGGER.info("Detected files to delete: {}", paths);
			JFiles.delete(paths);
			paths = JFiles.list(dirPath);
			if (paths.size() > 0) {
				LOGGER.info("Some files have not been deleted: {}", paths);
			}
		} else {
			Files.createDirectories(dirPath);
			LOGGER.info("Folder {} has been successfully created", dirPath);
		}
	}

	interface Test {

		void execute(final Spreadsheet spreadsheet) throws IOException;
	}

	private static class TestHeroes implements Test {

		private static final String SHEET_NAME = "Heroes";
		private static final String[] HEADER = new String[] { "Hero", "Origin", "Release Date", "Rarity", "HP" };

		@Override
		public void execute(final Spreadsheet spreadsheet) throws IOException {
			final Sheet sheet = spreadsheet.addSheet(SHEET_NAME);
			final SSCellFormat headerFormat = new SSCellFormat(SSFont.BOLD_ARIAL_10, SSColor.GREEN);
			sheet.setRow(0, HEADER, headerFormat);
			sheet.setRow(1, new Object[] { "Ike", "Fire Emblem: Path of Radiance", null, 5, 18 });
			sheet.setAutoSize();
			spreadsheet.write();
		}
	}

	private static class TestAppendRow implements Test {

		@Override
		public void execute(final Spreadsheet spreadsheet) throws IOException {
			if (spreadsheet.getSheetNames().size() == 0) {
				throw new IllegalStateException("No sheet to edit");
			}
			final Sheet sheet = spreadsheet.getSheet(TestHeroes.SHEET_NAME);
			final int rows = sheet.getRows();
			sheet.setRow(rows, new Object[] { "Raven", "Fire Emblem: The Blazing Blade", null, "4-5", 19 });
			sheet.setAutoSize();
			spreadsheet.write();
		}
	}

	private static class TestNewSheet implements Test {

		@Override
		public void execute(final Spreadsheet spreadsheet) throws IOException {
			final String newSheet = "New Sheet";
			final List<String> sheetNames = spreadsheet.getSheetNames();
			if (sheetNames.contains(newSheet)) {
				throw new IllegalStateException("Sheet with name '" + newSheet + "' already exist!");
			}
			final Sheet sheet = spreadsheet.addSheet("New Sheet");
			final int rows = MoreInts.newRandom(1, 10);
			for (int r = 0; r < rows; r++) {
				final int cols = MoreInts.newRandom(1, 12);
				final Object[] row = new Object[cols];
				for (int c = 0; c < cols; c++) {
					row[c] = MoreInts.newRandom(0, 1_000_000);
				}
				sheet.setRow(r, row);
			}
			spreadsheet.write();
		}
	}

	private static void create(final String outputDir, final String fileName, final API api, final Test test) throws IOException {
		final File file = Paths.get(outputDir, fileName).toFile();
		try (final Spreadsheet spreadsheet = Spreadsheets.create(file, api)) {
			test.execute(spreadsheet);
			LOGGER.info("File {} has been successfully created with API {}", fileName, api.name());
		}
	}

	private static void edit(final String outputDir, final String file, final API api, final Test test) throws IOException {
		final File in = Paths.get(outputDir, file).toFile();
		try (final Spreadsheet spreadsheet = Spreadsheets.edit(in, api)) {
			LOGGER.info("Editing file {} with API {}", in.getName(), api);
			test.execute(spreadsheet);
		}
	}

	private static void createFromTemplate(final String outputDir, final String fileIn, final String fileOut, final API api, final Test test)
			throws IOException {
		final File in = Paths.get(outputDir, fileIn).toFile();
		final File out = Paths.get(outputDir, fileOut).toFile();
		com.google.common.io.Files.copy(in, out);
		try (final Spreadsheet spreadsheet = Spreadsheets.edit(out, api)) {
			LOGGER.info("Editing file {} with API {}, created new file {}", in.getName(), api, out.getName());
			test.execute(spreadsheet);
		}
	}

	private static void openTryingEdit(final String outputDir, final API api, final String fileName) throws IOException {
		final File file = Paths.get(outputDir, fileName).toFile();
		try (final Spreadsheet spreadsheet = Spreadsheets.open(file, api)) {
			LOGGER.info("Opening file {} with API {}", file.getName(), api);
			final List<String> sheetNames = spreadsheet.getSheetNames();
			LOGGER.info("Sheet names of file {}: {}", file.getName(), sheetNames);
			final Sheet sheet = spreadsheet.getSheet(sheetNames.get(0));
			final int rows = sheet.getRows();
			// operazione non permessa su read-only file:
			sheet.setRow(rows + 3, new String[] { "2D", "3D", "4D" });
			// operazione non permessa su read-only file:
			spreadsheet.write();
		} catch (final IllegalStateException ex) {
			LOGGER.warn(ex.getMessage());
		}
	}
}
