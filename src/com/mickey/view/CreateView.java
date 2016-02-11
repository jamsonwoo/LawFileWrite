package com.mickey.view;

import javax.swing.JScrollPane;
import javax.swing.JTextArea;

public class CreateView {

	// 生成日志文本域
	public JScrollPane getJTextArea(JTextArea jtaLog) {
		if (jtaLog == null) {
			jtaLog = new JTextArea();
		}
		jtaLog.setLineWrap(true); // 激活自动换行功能
		jtaLog.setWrapStyleWord(true); // 激活断行不断字功能
		JScrollPane sp = new JScrollPane(jtaLog);
		sp.setBounds(5, 45, 650, 400);

		// 分别设置水平和垂直滚动条自动出现
		sp.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		sp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);

		// 分别设置水平和垂直滚动条总是出现
		// sp.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
		// sp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);

		// 分别设置水平和垂直滚动条总是隐藏
		// sp.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
		// sp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);

		return sp;
	}

	// 生成按钮
	/*
	 * public JButton getButton(JButton btn, String btnName) {
	 * 
	 * btn = new JButton();
	 * 
	 * btn.setText(btnName); System.out.println(btn); return btn; }
	 */
}
