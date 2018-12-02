package cz.gattserver.tulaci.calendar;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * https://www.programcreek.com/java-api-examples/?class=org.apache.poi.ss.usermodel.CellStyle&method=setFillForegroundColor
 * http://www.dominotricks.com/?p=115
 * http://viralpatel.net/blogs/java-read-write-excel-file-apache-poi/
 * https://www.callicoder.com/java-write-excel-file-apache-poi/
 * 
 * @author gattaka
 *
 */
public class Main {

	public static void main(String[] args) throws IOException {

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
		List<String> labelsFileLines = Files.readAllLines(Paths.get(prefix, labelsFileName));
		System.out.println("Budu brát data ze souboru: \t" + labelsFileName);

		String svatkyFileName = files.get(2);
		List<String> svatkyFileLines = Files.readAllLines(Paths.get(prefix, svatkyFileName));
		System.out.println("Budu brát svátky ze souboru: \t" + svatkyFileName);

		Map<String, String> svatkyMap = new HashMap<>();
		for (String svatek : svatkyFileLines) {
			String[] svatekData = svatek.split("\t");
			if (svatekData.length < 2) {
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
			if (birthdayData.length < 2) {
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

		File file = new File("Tuláci kalendář " + year + ".xlsx");

		try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOutputStream = new FileOutputStream(file)) {

			String[] sheetNames = new String[] { "První list", "Leden", "Únor", "Březen", "Duben", "Květen", "Červen",
					"Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec" };

			for (int i = 0; i < sheetNames.length; i++) {
				String sheetName = sheetNames[i];
				Sheet sheet = workbook.createSheet(sheetName);

				for (int c = 0; c < 7; c++)
					sheet.setColumnWidth(c, 3000);

				if (i == 0)
					continue;

				createMonthLine(sheet, sheetName, labelsFileLines.get(i - 1));

				LocalDate localDate = LocalDate.of(year, i, 1);
				int rowIndex = 23;
				Row dayRow = sheet.createRow(rowIndex);
				Row svatekRow = sheet.createRow(rowIndex + 1);
				Row birthdayRow = sheet.createRow(rowIndex + 2);
				while (i == localDate.getMonthValue()) {
					createDayCell(sheet, dayRow, localDate);
					createSvatekCell(sheet, svatekRow, localDate, svatkyMap);
					createBirthdayCell(sheet, birthdayRow, localDate, birthdaysMap);
					localDate = localDate.plusDays(1);
					// je další pondělí a jsem stále v tom stejném měsíci?
					if (localDate.getDayOfWeek().getValue() == 1 && i == localDate.getMonthValue()) {
						rowIndex += 3;
						dayRow = sheet.createRow(rowIndex);
						svatekRow = sheet.createRow(rowIndex + 1);
						birthdayRow = sheet.createRow(rowIndex + 2);
					}
				}

			}

			System.out.println("Zapisuji kalendář do souboru: \t" + file.getAbsolutePath());
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		}

	}

	private static void createBirthdayCell(Sheet sheet, Row birthdayRow, LocalDate localDate,
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

	private static void createSvatekCell(Sheet sheet, Row svatekRow, LocalDate localDate,
			Map<String, String> svatkyMap) {
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

	private static void createDayCell(Sheet sheet, Row dayRow, LocalDate localDate) {
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

	private static void createMonthLine(Sheet sheet, String sheetName, String quote) {
		Row row = sheet.createRow(0);
		createMonthName(sheet, row, sheetName);
		createMonthQuote(sheet, row, quote);
	}

	private static void createMonthName(Sheet sheet, Row row, String monthName) {
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

	private static void createMonthQuote(Sheet sheet, Row row, String quote) {
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

	private static void writeErrorSvatky(String errorLine) {
		showError("Řádek svátku '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -datum-tabulátor-text-\n" + "\tNapříklad: 17.1.\tDrahoslav");
	}

	private static void writeErrorBirthdays(String errorLine) {
		showError("Řádek narozenin '" + errorLine + "' má nevyhovující formát\n"
				+ "\tVyžaduji formát: -jméno-tabulátor-datum-\n" + "\tNapříklad: Vašek B.\t6.6.2008");
	}

	private static void showError(String msg) {
		JFrame frame = new JFrame();
		JOptionPane.showMessageDialog(frame, msg, "Chyba", JOptionPane.ERROR_MESSAGE);
		frame.dispose();
	}

}
