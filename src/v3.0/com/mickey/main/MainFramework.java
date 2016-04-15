package com.mickey.main;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Image;
import java.io.IOException;

import javax.imageio.ImageIO;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.UIManager;

import com.mickey.util.LFWConst;
import com.mickey.view.ActionLinster;
import com.mickey.view.CreateView;

public class MainFramework extends JFrame {

	private static final long serialVersionUID = 1L;

	private JTextArea jtaLog;

	private JTextField txtModelFile;
	private JTextField txtDataFile;
	private JTextField txtOutPath;

	private JButton btnModelFile;

	private JButton btnDataFile;
	private JButton btnOutPath;
	private JButton btnCreateFile;
	private JButton btnVersion;
	private JButton btnExit;

	private JRadioButton rdoAll;
	private JRadioButton rdoSingle;
	private JProgressBar progressBar;

	public static void main(String[] args) {
		try {
			MainFramework mainWindow = new MainFramework();
			mainWindow.setTitle(LFWConst.VERSION);
			mainWindow.InitMainWindow(mainWindow);
			mainWindow.setVisible(true);
		} catch (Exception e) {
		}
	}

	private void InitMainWindow(MainFramework mainWindow) throws IOException {
		UIManager.put("TextField.inactiveForeground", new Color(0, 0, 0));
		CreateView cv = new CreateView();

		// 设置大小
		setSize(720, 500);
		setLocationRelativeTo(null);
		setResizable(false);
		setDefaultCloseOperation(EXIT_ON_CLOSE);
		// mainWindow.dispatchEvent(new
		// WindowEvent(mainWindow,WindowEvent.WINDOW_CLOSING) );
		Image imgae = ImageIO.read(mainWindow.getClass().getResource("/com/mickey/resource/logo.gif"));
		setIconImage(imgae);

		JPanel contentPane;
		contentPane = new JPanel();
		setContentPane(contentPane);
		contentPane.setLayout(new BoxLayout(contentPane, BoxLayout.X_AXIS));

		contentPane.add(new JLabel("    "));

		JPanel panel = new JPanel(new BorderLayout(0, 0));
		contentPane.add(panel);

		JPanel panel_1 = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 5));
		panel.add(panel_1, BorderLayout.NORTH);

		JLabel label_1 = new JLabel("请选择模板文件：");
		panel_1.add(label_1);

		txtModelFile = new JTextField(LFWConst.MODEL_TEXT);
		// txtModelFile.setText("F:\\test\\结案报告表.doc");
		txtModelFile.setEnabled(false);
		// UIManager.put("TextField.inactiveForeground", Color.BLACK);
		panel_1.add(txtModelFile);
		txtModelFile.setColumns(40);

		btnModelFile = new JButton("选择模板");
		panel_1.add(btnModelFile);
		JPanel panel_2 = new JPanel(new BorderLayout(0, 0));
		panel.add(panel_2, BorderLayout.CENTER);

		JPanel panel_3 = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 5));
		panel_2.add(panel_3, BorderLayout.NORTH);

		JLabel label_2 = new JLabel("请选择数据文件：");
		panel_3.add(label_2);

		txtDataFile = new JTextField(LFWConst.DATA_TEXT);
		// txtDataFile.setText("F:\\test\\名单.xls");
		txtDataFile.setEnabled(false);
		// UIManager.put("TextField.inactiveForeground", Color.BLACK);
		// UIManager.put("TextField.font",Color.BLACK);
		panel_3.add(txtDataFile);
		txtDataFile.setColumns(40);

		btnDataFile = new JButton("选择数据");
		panel_3.add(btnDataFile);

		JPanel panel_4 = new JPanel(new BorderLayout(0, 0));
		panel_2.add(panel_4, BorderLayout.CENTER);

		JPanel panel_7 = new JPanel(new FlowLayout(FlowLayout.CENTER, 5, 5));
		panel_4.add(panel_7, BorderLayout.NORTH);
		JLabel label_4 = new JLabel("请选择输出路径：");
		panel_7.add(label_4);

		txtOutPath = new JTextField();
		// txtOutPath.setText("F:\\test\\");
		txtOutPath.setEnabled(false);
		panel_7.add(txtOutPath);
		txtOutPath.setColumns(40);

		btnOutPath = new JButton("选择路径");
		panel_7.add(btnOutPath);

		JPanel panel_5 = new JPanel();
		panel_4.add(panel_5, BorderLayout.CENTER);
		panel_5.setLayout(new BorderLayout(0, 0));

		JPanel panel_11 = new JPanel();
		panel_5.add(panel_11, BorderLayout.NORTH);
		panel_11.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));

		JLabel label_3 = new JLabel("请选择生成类型：");
		panel_11.add(label_3);

		rdoAll = new JRadioButton("生成整合文件");
		rdoAll.setSelected(true);
		panel_11.add(rdoAll);

		rdoSingle = new JRadioButton("生成单个文件");
		panel_11.add(rdoSingle);

		ButtonGroup group = new ButtonGroup();
		group.add(rdoAll);
		group.add(rdoSingle);

		JLabel lblNewLabel_1 = new JLabel("　       　");
		panel_11.add(lblNewLabel_1);

		btnCreateFile = new JButton("生成文档");
		panel_11.add(btnCreateFile);

		btnVersion = new JButton("版权声明");

		panel_11.add(btnVersion);

		btnExit = new JButton(" 退      出  ");
		panel_11.add(btnExit);
		
		JPanel panel_8 = new JPanel();
		panel_8.setLayout(new BorderLayout());
		panel_5.add(panel_8, BorderLayout.CENTER);
		jtaLog = new JTextArea();
		panel_8.add(cv.getJTextArea(jtaLog), BorderLayout.CENTER);

		JPanel panel_6 = new JPanel();
		panel_2.add(panel_6, BorderLayout.SOUTH);

		progressBar = new JProgressBar();
		progressBar.setOrientation(JProgressBar.HORIZONTAL);
		progressBar.setMinimum(0);
		progressBar.setMaximum(100);
		progressBar.setValue(0);
		progressBar.setStringPainted(true);
		progressBar.setPreferredSize(new Dimension(688, 24));
		progressBar.setBorderPainted(true);
		panel_6.add(progressBar);

		JLabel lblNewLabel = new JLabel("    ");
		contentPane.add(lblNewLabel);

		setActionListener(mainWindow);
	}

	private void setActionListener(MainFramework mainWindow) {
		ActionLinster al = new ActionLinster();
		al.setBtnVersionListener(btnVersion, mainWindow);
		al.setBtnExitListener(btnExit, mainWindow);

		al.setBtnModelFileListener(btnModelFile, mainWindow);
		al.setBtnDataFileListener(btnDataFile, mainWindow);
		al.setBtnOutPathListener(btnOutPath, mainWindow);

		al.setBtnCreateFileListener(btnCreateFile, mainWindow);
		
		al.setWindowListener(mainWindow);
	}

	public JTextArea getJtaLog() {
		return jtaLog;
	}

	public JTextField getTxtModelFile() {
		return txtModelFile;
	}

	public JTextField getTxtDataFile() {
		return txtDataFile;
	}

	public JTextField getTxtOutPath() {
		return txtOutPath;
	}

	public JButton getBtnModelFile() {
		return btnModelFile;
	}

	public JButton getBtnDataFile() {
		return btnDataFile;
	}

	public JButton getBtnOutPath() {
		return btnOutPath;
	}

	public JButton getBtnCreateFile() {
		return btnCreateFile;
	}

	public JButton getBtnVersion() {
		return btnVersion;
	}

	public JButton getBtnExit() {
		return btnExit;
	}

	public JRadioButton getRdoAll() {
		return rdoAll;
	}

	public JRadioButton getRdoSingle() {
		return rdoSingle;
	}

	public JProgressBar getProgressBar() {
		return progressBar;
	}

}