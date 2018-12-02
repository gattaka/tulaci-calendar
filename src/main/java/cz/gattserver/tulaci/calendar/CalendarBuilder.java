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
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static cz.gattserver.tulaci.calendar.GFXLogger.*;

public class CalendarBuilder {

	public CalendarBuilder() {
	}

	public void build() throws IOException {
		String prefix = "./data/";
		List<String> files = Files.readAllLines(Paths.get(prefix, "data.txt"));

		if (files.size() < 5) {
			showError(
					"Vyžaduji parametry: \n\t rok \n\t název souboru s popisky \n\t název souboru se svátky \n\t název souboru s narozeninami ");
			return;
		}

		int year;
		try {
			year = Integer.parseInt(files.get(0));
		} catch (NumberFormatException e) {
			showError("Rok má špatný formát: '" + files.get(0) + "', musí být celé číslo");
			return;
		}

		System.out.println("Generuji kalendář pro rok: \t" + year);

		String labelsFileName = files.get(1);
		Path labelsFilePath = Paths.get(prefix, labelsFileName);
		if (!Files.exists(labelsFilePath)) {
			showError("Soubor " + labelsFilePath.toString() + " neexistuje");
			return;
		}
		List<String> labelsFileLines = Files.readAllLines(labelsFilePath);
		System.out.println("Budu brát data ze souboru: \t" + labelsFileName);

		String svatkyFileName = files.get(2);
		List<String> svatkyFileLines = Files.readAllLines(Paths.get(prefix, svatkyFileName));
		System.out.println("Budu brát svátky ze souboru: \t" + svatkyFileName);

		Map<String, String> svatkyMap = new HashMap<>();
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

		Map<LocalDate, String> birthdaysMap = new HashMap<>();
		for (String birthday : birthdaysFileLines) {
			String[] birthdayData = birthday.split("\t");
			if (birthdayData.length != 2) {
				writeErrorBirthdays(birthday);
				return;
			}
			try {
				LocalDate date = LocalDate.parse(birthdayData[1], DateTimeFormatter.ofPattern("d.M.uuuu"));
				String label = birthdayData[0] + "(" + (year - date.getYear()) + ")";
				birthdaysMap.put(date.withYear(year), label);
			} catch (DateTimeParseException e) {
				writeErrorBirthdays(birthday);
				return;
			}
		}

		String fotoFileName = files.get(4);
		List<String> fotoFileLines = Files.readAllLines(Paths.get(prefix, fotoFileName));
		if (fotoFileLines.size() != 14) {
			showError("Chyba souboru '" + fotoFileName + "' očekávám následující obsah:\n"
					+ "\t jméno souboru fotky na první stránku\n"
					+ "\t jméno souboru fotky -tabulátor- popisek pro leden\n"
					+ "\t jméno souboru fotky -tabulátor- popisek pro únor\n" + "\t ...\n"
					+ "\t jméno souboru fotky -tabulátor- popisek pro prosinec\n"
					+ "\t jméno souboru fotky na poslední stránku\n");
			return;
		}

		System.out.println("Budu brát narozky ze souboru: \t" + fotoFileName);

		File file = new File("Tuláci kalendář " + year + ".xlsx");

		try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOutputStream = new FileOutputStream(file)) {

			String[] sheetNames = new String[] { "První list", "Leden", "Únor", "Březen", "Duben", "Květen", "Červen",
					"Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec" };

			for (int month = 0; month < sheetNames.length; month++) {
				String sheetName = sheetNames[month];
				Sheet sheet = workbook.createSheet(sheetName);

				for (int c = 0; c < 7; c++)
					sheet.setColumnWidth(c, 3000);

				if (month == 0)
					continue;

				// popisek měsíce + hláška
				createMonthLine(sheet, sheetName, labelsFileLines.get(month - 1));

				// fotka
				String fileLine = fotoFileLines.get(month + 1);
				String[] fileInfo = fileLine.split("\t");
				if (fileInfo.length != 2) {
					writeErrorFoto(fileLine);
					return;
				}

				Path photoPath = Paths.get(prefix, fileInfo[0]);
				if (!Files.exists(photoPath)) {
					showError("Soubor " + photoPath.toString() + " neexistuje");
					return;
				}

				final InputStream stream = Files.newInputStream(photoPath);
				final CreationHelper helper = workbook.getCreationHelper();
				final Drawing<?> drawing = sheet.createDrawingPatriarch();

				final ClientAnchor anchor = helper.createClientAnchor();
				anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

				final int pictureIndex = workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);

				anchor.setCol1(0);
				anchor.setRow1(1); // same row is okay
				anchor.setRow2(1);
				anchor.setCol2(0);
				final Picture pict = drawing.createPicture(anchor, pictureIndex);
				// double w = 800 / pict.getImageDimension().getWidth();
				// double h = 600 / pict.getImageDimension().getHeight();
				// pict.resize(w, h);

				// popisek fotky
				createPhotoLabel(sheet, fileInfo[1]);

				// dny, narozeniny a svátky
				createDaysTable(sheet, year, month, svatkyMap, birthdaysMap);
			}

			System.out.println("Zapisuji kalendář do souboru: \t" + file.getAbsolutePath());
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		}
	}

	private void createDaysTable(Sheet sheet, int year, int month, Map<String, String> svatkyMap,
			Map<LocalDate, String> birthdaysMap) {
		LocalDate localDate = LocalDate.of(year, month, 1);
		int rowIndex = 23;
		Row dayRow = sheet.createRow(rowIndex);
		Row svatekRow = sheet.createRow(rowIndex + 1);
		Row birthdayRow = sheet.createRow(rowIndex + 2);
		while (month == localDate.getMonthValue()) {
			createDayCell(sheet, dayRow, localDate);
			createSvatekCell(sheet, svatekRow, localDate, svatkyMap);
			createBirthdayCell(sheet, birthdayRow, localDate, birthdaysMap);
			localDate = localDate.plusDays(1);
			// je další pondělí a jsem stále v tom stejném měsíci?
			if (localDate.getDayOfWeek().getValue() == 1 && month == localDate.getMonthValue()) {
				rowIndex += 3;
				dayRow = sheet.createRow(rowIndex);
				svatekRow = sheet.createRow(rowIndex + 1);
				birthdayRow = sheet.createRow(rowIndex + 2);
			}
		}
	}

	private void createPhotoLabel(Sheet sheet, String label) {
		Row photoLabelRow = sheet.createRow(22);
		Cell cell = photoLabelRow.createCell(0);
		sheet.addMergedRegion(new CellRangeAddress(22, 22, 0, 6));
		cell.setCellValue(label);

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 8);
		style.setFont(font);

		cell.setCellStyle(style);
	}

	private void createBirthdayCell(Sheet sheet, Row birthdayRow, LocalDate localDate,
			Map<LocalDate, String> birthdaysMap) {
		Cell cell = birthdayRow.createCell(localDate.getDayOfWeek().getValue() - 1);
		cell.setCellValue(birthdaysMap.get(localDate));

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 8);
		font.setColor(HSSFColorPredefined.DARK_RED.getIndex());

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		cell.setCellStyle(style);
	}

	private void createSvatekCell(Sheet sheet, Row svatekRow, LocalDate localDate, Map<String, String> svatkyMap) {
		Cell cell = svatekRow.createCell(localDate.getDayOfWeek().getValue() - 1);
		cell.setCellValue(svatkyMap.get(localDate.getDayOfMonth() + "." + localDate.getMonthValue() + "."));

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 8);

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
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 6));
		cell.setCellValue(quote);

		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Castanet CE");
		font.setBold(false);
		font.setFontHeightInPoints((short) 14);
		style.setFont(font);

		cell.setCellStyle(style);
	}

	private void writeErrorSvatky(String errorLine) {
		showError("Řádek svátku '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -datum-tabulátor-text-\n" + "\tNapříklad: 17.1.\tDrahoslav");
	}

	private void writeErrorBirthdays(String errorLine) {
		showError("Řádek narozenin '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -jméno-tabulátor-datum-\n" + "\tNapříklad: Vašek B.\t6.6.2008");
	}

	private void writeErrorFoto(String errorLine) {
		showError("Řádek fotky '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -soubor.přípona-tabulátor-název akce-\n"
				+ "\tNapříklad: foto1.jpg\tVýprava na Sněžku");
	}

}
