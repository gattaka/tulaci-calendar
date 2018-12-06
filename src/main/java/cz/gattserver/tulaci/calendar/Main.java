package cz.gattserver.tulaci.calendar;

/**
 * https://www.programcreek.com/java-api-examples/?class=org.apache.poi.ss.usermodel.CellStyle&method=setFillForegroundColor
 * http://www.dominotricks.com/?p=115 http://viralpatel.net/blogs/java-read-write-excel-file-apache-poi/
 * https://www.callicoder.com/java-write-excel-file-apache-poi/
 * 
 * @author gattaka
 *
 */
public class Main {

	public static void main(String[] args) throws Exception {

		CalendarBuilder calendarBuilder = new CalendarBuilder();
		try {
			calendarBuilder.build();
			GFXLogger.showSuccess("Generování kalendáře dopadlo úspěšně");
		} catch (Exception e) {
			GFXLogger.showError(e.getMessage());
			throw e;
		}

	}

}
