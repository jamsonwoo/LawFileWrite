package com.mickey.view;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.mickey.core.HandleFile;
import com.mickey.main.MainFramework;
import com.mickey.util.LFWConst;
import com.mickey.util.LFWStringUtil;
import com.mickey.util.LFWUtil;

public class ActionLinster {
	public void setBtnExitListener(JButton btnExit, final MainFramework mainWindow) {
		btnExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				int exi = JOptionPane.showConfirmDialog(null, "Ҫ�˳��ó�����", "������ʾ", JOptionPane.YES_NO_OPTION,
						JOptionPane.QUESTION_MESSAGE);
				if (exi == JOptionPane.YES_OPTION) {
					mainWindow.dispose();
					System.exit(0);
				} else {
					return;
				}
			}
		});
	}

	public void setBtnVersionListener(JButton btnVersion, MainFramework mainWindow) {
		btnVersion.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JOptionPane.showMessageDialog(null, "������� ����������ʦ������ ��Ȩ����", "��Ȩ����", JOptionPane.WARNING_MESSAGE);
			}
		});
	}

	public void setBtnModelFileListener(JButton btn, final MainFramework mainWindow) {
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Office Word Doc �ļ�", "doc");
				jfc.setFileFilter(filter);// �����ļ���׺������
				jfc.showDialog(new JLabel(), "ѡ��");

				File f = jfc.getSelectedFile();// fΪѡ�񵽵�Ŀ¼
				if (f != null && f.exists())
					mainWindow.getTxtModelFile().setText(f.getAbsolutePath());
			}
		});
	}

	public void setBtnDataFileListener(JButton btn, final MainFramework mainWindow) {
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Office Excel Xls �ļ�", "xls");
				jfc.setFileFilter(filter);// �����ļ���׺������
				jfc.showDialog(new JLabel(), "ѡ��");
				File f = jfc.getSelectedFile();// fΪѡ�񵽵�Ŀ¼
				mainWindow.getTxtDataFile().setText(f.getAbsolutePath());

				String path = f.getAbsolutePath();
				mainWindow.getTxtOutPath().setText(LFWStringUtil.getFilePath(path));
			}
		});
	}

	public void setBtnOutPathListener(JButton btn, final MainFramework mainWindow) {
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				jfc.showDialog(new JLabel(), "ѡ��");
				File f = jfc.getSelectedFile();// fΪѡ�񵽵�Ŀ¼
				mainWindow.getTxtOutPath().setText(f.getAbsolutePath());
			}
		});
	}

	public void setBtnCreateFileListener(final JButton btn, final MainFramework mainWindow) {
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				btn.setEnabled(false);
				btn.setText("������...");

				String modelFile = mainWindow.getTxtModelFile().getText();
				String dataFile = mainWindow.getTxtDataFile().getText();
				String outPath = mainWindow.getTxtOutPath().getText();

				if (modelFile.equals(LFWConst.MODEL_TEXT)) {
					LFWUtil.Alert("��ѡ��ģ���ļ���", JOptionPane.ERROR_MESSAGE);
					return;
				}
				if (dataFile.equals(LFWConst.DATA_TEXT)) {
					LFWUtil.Alert("��ѡ�������ļ���", JOptionPane.ERROR_MESSAGE);
					return;
				}
				if (outPath.equals("")) {// ||(!(new File(outPath).exists()))
					LFWUtil.Alert("��ѡ������ļ��У�", JOptionPane.ERROR_MESSAGE);
					return;
				}

				// handle(String dataPath, String modelFileName, String
				// srcFileName, String createFileTypy)
				char createFileType = mainWindow.getRdoAll().isSelected() ? 'a' : 's';

				HandleFile hf = new HandleFile(modelFile, dataFile, outPath, createFileType, mainWindow);
				new Thread(hf).start();
			}
		});
	}

	public void setWindowListener(final MainFramework mainWindow) {
		mainWindow.addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent e) {

				/*int exi = JOptionPane.showConfirmDialog(null, "Ҫ�˳��ó�����", "������ʾ", JOptionPane.YES_NO_OPTION,
						JOptionPane.QUESTION_MESSAGE);
				System.out.println(exi);
				if (exi == JOptionPane.YES_OPTION) {*/
					mainWindow.dispose();
					System.exit(0);
				/*} else {
					System.out.println("test");
				}*/
			}
		});
	}
}
