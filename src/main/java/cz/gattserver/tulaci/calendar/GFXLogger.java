package cz.gattserver.tulaci.calendar;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

public class GFXLogger {

	public static void showError(String msg) {
		JFrame frame = new JFrame();
		JOptionPane.showMessageDialog(frame, msg, "Chyba", JOptionPane.ERROR_MESSAGE);
		frame.dispose();
	}

}
