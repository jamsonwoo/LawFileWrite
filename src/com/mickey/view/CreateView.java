package com.mickey.view;

import javax.swing.JScrollPane;
import javax.swing.JTextArea;

public class CreateView {

	// ������־�ı���
	public JScrollPane getJTextArea(JTextArea jtaLog) {
		if (jtaLog == null) {
			jtaLog = new JTextArea();
		}
		jtaLog.setLineWrap(true); // �����Զ����й���
		jtaLog.setWrapStyleWord(true); // ������в����ֹ���
		JScrollPane sp = new JScrollPane(jtaLog);
		sp.setBounds(5, 45, 650, 400);

		// �ֱ�����ˮƽ�ʹ�ֱ�������Զ�����
		sp.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		sp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);

		// �ֱ�����ˮƽ�ʹ�ֱ���������ǳ���
		// sp.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
		// sp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);

		// �ֱ�����ˮƽ�ʹ�ֱ��������������
		// sp.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
		// sp.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_NEVER);

		return sp;
	}

	// ���ɰ�ť
	/*
	 * public JButton getButton(JButton btn, String btnName) {
	 * 
	 * btn = new JButton();
	 * 
	 * btn.setText(btnName); System.out.println(btn); return btn; }
	 */
}
