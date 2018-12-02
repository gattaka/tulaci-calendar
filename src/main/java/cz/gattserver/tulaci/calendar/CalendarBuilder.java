package cz.gattserver.tulaci.calendar;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CalendarBuilder {

	private static final String[] sheetNames = new String[] { "První list", "Leden", "Únor", "Březen", "Duben",
			"Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec", "Poslední list" };

	private static final String prefix = "./data/";

	private int year;
	private List<String> fotoFileLines;
	private List<String> labelsFileLines;

	private Map<String, String> svatkyMap;
	private Map<LocalDate, String> birthdaysMap;

	private static class BirthdayEntry {
		int day;
		String name;

		public BirthdayEntry(int day, String name) {
			this.day = day;
			this.name = name;
		}
	}

	public CalendarBuilder() {
	}

	public void build() throws IOException {

		Path dataFilePath = Paths.get(prefix, "data.txt");
		if (!Files.exists(dataFilePath))
			throw new IllegalStateException("Soubor " + dataFilePath.toString() + " neexistuje");
		List<String> files = Files.readAllLines(dataFilePath);

		if (files.size() < 5)
			throw new IllegalStateException(
					"Vyžaduji parametry: \n\t rok \n\t název souboru s popisky \n\t název souboru se svátky \n\t název souboru s narozeninami ");

		try {
			year = Integer.parseInt(files.get(0));
		} catch (NumberFormatException e) {
			throw new IllegalStateException("Rok má špatný formát: '" + files.get(0) + "', musí být celé číslo");
		}

		System.out.println("Generuji kalendář pro rok: \t" + year);

		String labelsFileName = files.get(1);
		Path labelsFilePath = Paths.get(prefix, labelsFileName);
		if (!Files.exists(labelsFilePath)) {
			throw new IllegalStateException("Soubor " + labelsFilePath.toString() + " neexistuje");
		}
		labelsFileLines = Files.readAllLines(labelsFilePath);
		System.out.println("Budu brát data ze souboru: \t" + labelsFileName);

		String svatkyFileName = files.get(2);
		List<String> svatkyFileLines = Files.readAllLines(Paths.get(prefix, svatkyFileName));
		System.out.println("Budu brát svátky ze souboru: \t" + svatkyFileName);

		svatkyMap = new HashMap<>();
		for (String svatek : svatkyFileLines) {
			String[] svatekData = svatek.split("\t");
			if (svatekData.length != 2) {
				writeErrorSvatky(svatek);
				return;
			}
			if (!svatekData[0].matches("[1-3]?[0-9]\\.[1]?[0-9]\\.")) {
				writeErrorSvatky(svatek);
				return;
			}
			svatkyMap.put(svatekData[0], svatekData[1]);
		}

		String birthdaysFileName = files.get(3);
		List<String> birthdaysFileLines = Files.readAllLines(Paths.get(prefix, birthdaysFileName));
		System.out.println("Budu brát narozky ze souboru: \t" + birthdaysFileName);

		birthdaysMap = new HashMap<>();
		for (String birthday : birthdaysFileLines) {
			String[] birthdayData = birthday.split("\t");
			if (birthdayData.length != 2) {
				writeErrorBirthdays(birthday);
				return;
			}
			try {
				LocalDate date = LocalDate.parse(birthdayData[1], DateTimeFormatter.ofPattern("d.M.uuuu"));
				String label = birthdayData[0] + " (" + (year - date.getYear()) + ")";
				birthdaysMap.put(date.withYear(year), label);
			} catch (DateTimeParseException e) {
				writeErrorBirthdays(birthday);
				return;
			}
		}

		String fotoFileName = files.get(4);
		fotoFileLines = Files.readAllLines(Paths.get(prefix, fotoFileName));
		if (fotoFileLines.size() != 14) {
			throw new IllegalStateException("Chyba souboru '" + fotoFileName + "' očekávám následující obsah:\n"
					+ "\t jméno souboru fotky na první stránku\n"
					+ "\t jméno souboru fotky -tabulátor- popisek pro leden\n"
					+ "\t jméno souboru fotky -tabulátor- popisek pro únor\n" + "\t ...\n"
					+ "\t jméno souboru fotky -tabulátor- popisek pro prosinec\n"
					+ "\t jméno souboru fotky na poslední stránku\n");
		}

		System.out.println("Budu brát narozky ze souboru: \t" + fotoFileName);

		File file = new File("Tuláci kalendář " + year + ".xlsx");

		try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOutputStream = new FileOutputStream(file)) {

			for (int sheetNo = 0; sheetNo < sheetNames.length; sheetNo++) {
				Sheet sheet = workbook.createSheet(sheetNames[sheetNo]);

				for (int c = 0; c < 7; c++)
					sheet.setColumnWidth(c, 3000);

				switch (sheetNo) {
				case 0:
					createFrontSheet(workbook, sheet, sheetNo);
					break;
				case 13:
					createBackSheet(workbook, sheet, sheetNo);
					break;
				default:
					createMonthSheet(workbook, sheet, sheetNo);
					break;
				}

			}

			System.out.println("Zapisuji kalendář do souboru: \t" + file.getAbsolutePath());
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		}
	}

	private void createFrontSheet(Workbook workbook, Sheet sheet, int sheetNo) throws IOException {

		// fotka
		String fileLine = fotoFileLines.get(sheetNo);
		String[] fileInfo = fileLine.split("\t");
		if (fileInfo.length != 2)
			writeErrorFoto(fileLine);

		Path photoPath = Paths.get(prefix, fileInfo[0]);
		if (!Files.exists(photoPath))
			throw new IllegalStateException("Soubor " + photoPath.toString() + " neexistuje");

		final InputStream stream = Files.newInputStream(photoPath);
		final CreationHelper helper = workbook.getCreationHelper();
		final Drawing<?> drawing = sheet.createDrawingPatriarch();

		final ClientAnchor anchor = helper.createClientAnchor();
		anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

		final int pictureIndex = workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);

		anchor.setCol1(0);
		anchor.setCol2(7);
		anchor.setRow1(0);
		anchor.setRow2(39);
		drawing.createPicture(anchor, pictureIndex);
	}

	private void createBackSheet(Workbook workbook, Sheet sheet, int sheetNo) throws IOException {

		int line = 0;

		/*
		 * text 1
		 */

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 38);
		font.setColor(HSSFColorPredefined.BROWN.getIndex());

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		Row row = sheet.createRow(line);
		Cell cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line + 2, 0, 6));
		cell.setCellValue("Tulácký kalendář");

		line += 3;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line + 2, 0, 6));
		cell.setCellValue(year);

		line += 3;
		line++;

		/*
		 * text 2
		 */

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 22);
		font.setColor(HSSFColorPredefined.DARK_GREEN.getIndex());

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("Aktivity našeho oddílu jsou podporovány");

		line++;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("mladými ochránci přírody z prostředků");

		line++;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("MŠMT a MHMP.");

		line++;
		line++;

		// fotka
		String fileLine = fotoFileLines.get(sheetNo);
		String[] fileInfo = fileLine.split("\t");
		if (fileInfo.length != 2)
			writeErrorFoto(fileLine);

		Path photoPath = Paths.get(prefix, fileInfo[0]);
		if (!Files.exists(photoPath))
			throw new IllegalStateException("Soubor " + photoPath.toString() + " neexistuje");

		final InputStream stream = Files.newInputStream(photoPath);
		final CreationHelper helper = workbook.getCreationHelper();
		final Drawing<?> drawing = sheet.createDrawingPatriarch();

		final ClientAnchor anchor = helper.createClientAnchor();
		anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

		final int pictureIndex = workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);

		anchor.setCol1(0);
		anchor.setCol2(7);
		anchor.setRow1(line);
		anchor.setRow2(line + 10);
		drawing.createPicture(anchor, pictureIndex);

		line += 11;

		/*
		 * text 3
		 */

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 12);
		font.setColor(HSSFColorPredefined.DARK_GREEN.getIndex());

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("vydáno jako 95. publikace oddílového nakladatelství NAKOLENĚ");

		line++;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("neprodejný materiál pro členy a příznivce oddílu TULÁCI");

		line++;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("prosinec " + (year - 1) + ", vydání prvé, náklad závratný (25 ks)");

		line++;
		line++;

		/*
		 * text 4
		 */

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 12);
		font.setColor(HSSFColorPredefined.DARK_GREEN.getIndex());

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("Všechny fotografie pochází z fotoaparátů členů oddílu, jakákoliv podobnost");

		line++;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue(" s fotografiemi jiných autorů je čistě náhodná. ");

		line++;

		/*
		 * text 5
		 */

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 12);
		font.setColor(HSSFColorPredefined.DARK_RED.getIndex());

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("Neneseme odpovědnost za pohoršení při prohlížení kalendáře.");

		line++;
		line++;

		/*
		 * text 6
		 */

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 15);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("Kontakt na oddíl (působící v Praze 10)");

		line++;

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("Klára Adámková, tel.: 728 734 009, email: oddil@tulaci.eu");

		line++;
		line++;

		/*
		 * text 7
		 */

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 15);
		font.setColor(HSSFColorPredefined.DARK_RED.getIndex());

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);

		row = sheet.createRow(line);
		cell = row.createCell(0);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(line, line, 0, 6));
		cell.setCellValue("Vše o nás najdete na http://oddil.tulaci.eu");

	}

	private void createMonthSheet(Workbook workbook, Sheet sheet, int sheetNo) throws IOException {

		// popisek měsíce + hláška
		createMonthLine(sheet, sheetNames[sheetNo], labelsFileLines.get(sheetNo - 1));

		// fotka
		String fileLine = fotoFileLines.get(sheetNo);
		String[] fileInfo = fileLine.split("\t");
		if (fileInfo.length != 2)
			writeErrorFoto(fileLine);

		Path photoPath = Paths.get(prefix, fileInfo[0]);
		if (!Files.exists(photoPath))
			throw new IllegalStateException("Soubor " + photoPath.toString() + " neexistuje");

		final InputStream stream = Files.newInputStream(photoPath);
		final CreationHelper helper = workbook.getCreationHelper();
		final Drawing<?> drawing = sheet.createDrawingPatriarch();

		final ClientAnchor anchor = helper.createClientAnchor();
		anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

		final int pictureIndex = workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);

		anchor.setCol1(0);
		anchor.setCol2(10);
		anchor.setRow1(1);
		anchor.setRow2(22);
		drawing.createPicture(anchor, pictureIndex);

		// popisek fotky
		// createPhotoLabel(sheet, fileInfo[1]);

		// dny, narozeniny a svátky
		createDaysTable(sheet, sheetNo);
	}

	private void createDaysTable(Sheet sheet, int month) {

		List<BirthdayEntry> birthdays = new ArrayList<>();

		LocalDate localDate = LocalDate.of(year, month, 1);
		int rowStart = 23;
		int rowIndex = rowStart;
		Row dayRow = sheet.createRow(rowIndex);
		Row svatekRow = sheet.createRow(rowIndex + 2);
		while (month == localDate.getMonthValue()) {
			createDayCell(sheet, dayRow, localDate);
			createSvatekCell(sheet, svatekRow, localDate);
			String value = birthdaysMap.get(localDate);
			if (value != null)
				birthdays.add(new BirthdayEntry(localDate.getDayOfMonth(), value));
			localDate = localDate.plusDays(1);
			// je další pondělí a jsem stále v tom stejném měsíci?
			if (localDate.getDayOfWeek().getValue() == 1 && month == localDate.getMonthValue()) {
				rowIndex += 3;
				dayRow = sheet.createRow(rowIndex);
				svatekRow = sheet.createRow(rowIndex + 2);
			}
		}

		createBirthdayList(sheet, rowStart, birthdays);
	}

	private void createBirthdayList(Sheet sheet, int rowStart, List<BirthdayEntry> birthdays) {
		// Birthdays list
		Row headerRow = sheet.getRow(rowStart);
		Cell cell = headerRow.createCell(7);
		sheet.addMergedRegion(new CellRangeAddress(rowStart, rowStart, 7, 9));
		cell.setCellValue("Narozeniny");

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 12);
		style.setFont(font);

		cell.setCellStyle(style);

		sheet.setColumnWidth(7, 1000);

		int currentRowIndex = rowStart + 1;
		for (BirthdayEntry be : birthdays) {

			Row beRow = sheet.getRow(currentRowIndex);
			if (beRow == null)
				beRow = sheet.createRow(currentRowIndex);

			Cell beCell = beRow.createCell(7);
			beCell.setCellValue(be.day);

			style = sheet.getWorkbook().createCellStyle();
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setVerticalAlignment(VerticalAlignment.CENTER);

			font = sheet.getWorkbook().createFont();
			font.setFontName("Castanet CE");
			font.setBold(false);
			font.setColor(HSSFColorPredefined.DARK_RED.getIndex());
			font.setFontHeightInPoints((short) 8);
			style.setFont(font);

			beCell.setCellStyle(style);

			beCell = beRow.createCell(8);
			sheet.addMergedRegion(new CellRangeAddress(currentRowIndex, currentRowIndex, 8, 9));
			beCell.setCellValue(be.name);

			style = sheet.getWorkbook().createCellStyle();
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setVerticalAlignment(VerticalAlignment.CENTER);

			font = sheet.getWorkbook().createFont();
			font.setFontName("Castanet CE");
			font.setBold(false);
			font.setFontHeightInPoints((short) 8);
			style.setFont(font);

			beCell.setCellStyle(style);

			currentRowIndex++;
		}

		// Akce
		int headerOffset = 8;
		int headerRowIndex = rowStart + headerOffset;
		headerRow = sheet.getRow(headerRowIndex);
		if (headerRow == null)
			headerRow = sheet.createRow(headerRowIndex);

		cell = headerRow.createCell(7);
		sheet.addMergedRegion(new CellRangeAddress(headerRowIndex, headerRowIndex, 7, 9));
		cell.setCellValue("Nezapomeň!");

		style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setColor(HSSFColorPredefined.DARK_RED.getIndex());
		font.setFontHeightInPoints((short) 12);
		style.setFont(font);

		cell.setCellStyle(style);
	}

	private void createSvatekCell(Sheet sheet, Row svatekRow, LocalDate localDate) {
		Cell cell = svatekRow.createCell(localDate.getDayOfWeek().getValue() - 1);
		cell.setCellValue(svatkyMap.get(localDate.getDayOfMonth() + "." + localDate.getMonthValue() + "."));

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 7);

		if (localDate.getDayOfWeek().getValue() > 5)
			font.setColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		cell.setCellStyle(style);
	}

	private void createDayCell(Sheet sheet, Row dayRow, LocalDate localDate) {
		Cell cell = dayRow.createCell(localDate.getDayOfWeek().getValue() - 1);
		sheet.addMergedRegion(new CellRangeAddress(dayRow.getRowNum(), dayRow.getRowNum() + 1, cell.getColumnIndex(),
				cell.getColumnIndex()));
		cell.setCellValue(localDate.getDayOfMonth());

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 28);

		if (localDate.getDayOfWeek().getValue() > 5)
			font.setColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
		cell.setCellStyle(style);
	}

	private void createMonthLine(Sheet sheet, String sheetName, String quote) {
		Row row = sheet.createRow(0);
		createMonthName(sheet, row, sheetName);
		createMonthQuote(sheet, row, quote);
	}

	private void createMonthName(Sheet sheet, Row row, String monthName) {
		Cell cell = row.createCell(0);
		cell.setCellValue(monthName.toUpperCase());

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 32);
		font.setColor(HSSFColorPredefined.BLUE.getIndex());

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setFont(font);
		cell.setCellStyle(style);
	}

	private void createMonthQuote(Sheet sheet, Row row, String quote) {
		Cell cell = row.createCell(3);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 9));
		cell.setCellValue(quote);

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 10);
		style.setFont(font);

		cell.setCellStyle(style);
	}

	private void writeErrorSvatky(String errorLine) {
		throw new IllegalStateException("Řádek svátku '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -datum-tabulátor-text-\n" + "\tNapříklad: 17.1.\tDrahoslav");
	}

	private void writeErrorBirthdays(String errorLine) {
		throw new IllegalStateException("Řádek narozenin '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -jméno-tabulátor-datum-\n" + "\tNapříklad: Vašek B.\t6.6.2008");
	}

	private void writeErrorFoto(String errorLine) {
		throw new IllegalStateException("Řádek fotky '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -soubor.přípona-tabulátor-název akce-\n"
				+ "\tNapříklad: foto1.jpg\tVýprava na Sněžku");
	}

}
