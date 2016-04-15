package com.mickey.util;

import javax.swing.JTextArea;

public class LFWLog {
	public static void writeLog(JTextArea jtaLog, String log) {
		String oldLog = jtaLog.getText();
		StringBuffer newLog = new StringBuffer();

		newLog.append(log);
		newLog.append("\n");
		newLog.append(oldLog);
		jtaLog.setText(newLog.toString());
	}
}
