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
				int exi = JOptionPane.showConfirmDialog(null, "要退出该程序吗？", "友情提示", JOptionPane.YES_NO_OPTION,
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
				JOptionPane.showMessageDialog(null, "本软件由 广西东中律师事务所 版权所有", "版权声明", JOptionPane.WARNING_MESSAGE);
			}
		});
	}

	public void setBtnModelFileListener(JButton btn, final MainFramework mainWindow) {
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Office Word Doc 文件", "doc");
				jfc.setFileFilter(filter);// 设置文件后缀过滤器
				jfc.showDialog(new JLabel(), "选择");

				File f = jfc.getSelectedFile();// f为选择到的目录
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
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Office Excel Xls 文件", "xls");
				jfc.setFileFilter(filter);// 设置文件后缀过滤器
				jfc.showDialog(new JLabel(), "选择");
				File f = jfc.getSelectedFile();// f为选择到的目录
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
				jfc.showDialog(new JLabel(), "选择");
				File f = jfc.getSelectedFile();// f为选择到的目录
				mainWindow.getTxtOutPath().setText(f.getAbsolutePath());
			}
		});
	}

	public void setBtnCreateFileListener(final JButton btn, final MainFramework mainWindow) {
		btn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				btn.setEnabled(false);
				btn.setText("处理中...");

				String modelFile = mainWindow.getTxtModelFile().getText();
				String dataFile = mainWindow.getTxtDataFile().getText();
				String outPath = mainWindow.getTxtOutPath().getText();

				if (modelFile.equals(LFWConst.MODEL_TEXT)) {
					LFWUtil.Alert("请选择模板文件！", JOptionPane.ERROR_MESSAGE);
					return;
				}
				if (dataFile.equals(LFWConst.DATA_TEXT)) {
					LFWUtil.Alert("请选择数据文件！", JOptionPane.ERROR_MESSAGE);
					return;
				}
				if (outPath.equals("")) {// ||(!(new File(outPath).exists()))
					LFWUtil.Alert("请选择输出文件夹！", JOptionPane.ERROR_MESSAGE);
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

				/*int exi = JOptionPane.showConfirmDialog(null, "要退出该程序吗？", "友情提示", JOptionPane.YES_NO_OPTION,
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
