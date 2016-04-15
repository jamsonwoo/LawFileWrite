package com.mickey.util;

import javax.swing.JOptionPane;

public class LFWUtil {
	public static void Alert(String warnText, int type) {
		JOptionPane.showMessageDialog(null, warnText, "提示", type);
	}

	// 数字转字母 0-A 1-B 2-C。。。
	public static char toTag(int i) {
		return (char)(i+65);
	}
	
	public static void main(String[] args) {
		for (int i = 0; i < 26; i++) {
			System.out.println(toTag(i));
		}
	}
}
