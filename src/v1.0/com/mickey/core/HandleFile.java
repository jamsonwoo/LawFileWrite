package com.mickey.core;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.JTextArea;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.FieldsDocumentPart;
import org.apache.poi.hwpf.usermodel.Field;
import org.apache.poi.hwpf.usermodel.Fields;
import org.apache.poi.hwpf.usermodel.Range;

import com.mickey.main.MainFramework;
import com.mickey.util.LFWConst;
import com.mickey.util.LFWLog;
import com.mickey.util.LFWStringUtil;
import com.mickey.util.LFWUtil;

public class HandleFile extends Thread {

	private static int index = 1;

	// F:\test �᰸����� ����.xls a

	private List<Map<String, String>> readSrc(String dataFile) throws Exception {
		// System.out.println("���ڴ��������ļ���" + dataFile + "\n");
		LFWLog.writeLog(jtaLog, "���ڴ��������ļ���" + dataFile + "");
		File file = new File(dataFile);

		List<Map<String, String>> data = new ArrayList<Map<String, String>>();
		data = ExcelOperate.getData(file);
		return data;

	}

	/**
	 * ʵ�ֶ�word��ȡ���޸Ĳ���
	 * 
	 * @param filePath
	 *            wordģ��·��������
	 * @param map
	 *            ���������ݣ������ݿ��ȡ
	 */
	public void readwriteWord(String modelFile, String outPath, char createFileTypy, Map<String, String> map)
			throws Exception {
		// ��ȡwordģ��
		FileInputStream in = null;
		ByteArrayOutputStream ostream = new ByteArrayOutputStream();
		FileOutputStream out = null;

		String dirName = LFWStringUtil.getFileNameByFilePath(modelFile);
		String outFileName = dirName + LFWConst.SPLIT;

		if (createFileTypy == 'a') {
			dirName = LFWConst.FILE_TMP;
			outFileName = LFWConst.FILE_TMP + LFWConst.SPLIT + (index++);
		}
		// System.out.println("dirName=" + dirName);
		// System.out.println("outFileName=" + outFileName);
		// System.out.println("map=" + map);
		// System.out.println("modelFile=" + modelFile);

		try {
			File f = new File(modelFile);
			in = new FileInputStream(f);

			HWPFDocument hdt = null;
			hdt = new HWPFDocument(in);

			Fields fields = hdt.getFields();
			Iterator<Field> it = fields.getFields(FieldsDocumentPart.MAIN).iterator();
			while (it.hasNext()) {
				System.out.println(it.next().getType());
			}

			// ��ȡword�ı�����
			Range range = hdt.getRange();
			/*
			 * TableIterator tableIt = new TableIterator(range); // �����ĵ��еı��
			 * while (tableIt.hasNext()) { Table tb = (Table) tableIt.next(); //
			 * �����У�Ĭ�ϴ�0��ʼ for (int i = 0; i < tb.numRows(); i++) { TableRow tr =
			 * tb.getRow(i); // ֻ��ǰ8�У����ⲿ�� if (i >= 8) break; // �����У�Ĭ�ϴ�0��ʼ for
			 * (int j = 0; j < tr.numCells(); j++) { TableCell td =
			 * tr.getCell(j);// ȡ�õ�Ԫ�� // ȡ�õ�Ԫ������� for (int k = 0; k <
			 * td.numParagraphs(); k++) { Paragraph para = td.getParagraph(k);
			 * // String s = para.text(); // System.out.println(s); } } } }
			 */
			//

			// �滻�ı�����
			for (Map.Entry<String, String> entry : map.entrySet()) {
				range.replaceText(entry.getKey(), entry.getValue());
			}
			// System.out.println(range.text());

			File folder = new File(outPath + "\\" + dirName);
			if (!(folder.exists() && folder.isDirectory())) {
				folder.mkdirs();
			}
			String fn = folder + "\\" + outFileName + "_" + map.get("${1}") + LFWConst.FILE_TYPE;
			// System.out.println("���ɽ���ļ���" + fn + "\n");
			LFWLog.writeLog(jtaLog, "���ɽ���ļ���" + fn + "\n");
			out = new FileOutputStream(fn, false);
			hdt.write(ostream);
			out.write(ostream.toByteArray());

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			out.close();
			ostream.close();
		}
	}

	private static void mergeFiles(File tmpFile, String createFile) throws IOException {
		// System.out.println(tmpFile.getPath());
		// System.out.println(createFile.getPath());

		List<String> list = new ArrayList<String>();
		String[] children = tmpFile.list();
		// �ݹ�ɾ��Ŀ¼�е���Ŀ¼��
		for (int i = 0; i < children.length; i++) {
			// System.out.println(tmpFile+"\\"+children[i]);
			list.add(tmpFile + "\\" + children[i]);
		}

		MergeDoc.uniteDoc(list, createFile);

	}

	private static boolean deleteDir(File dir) {
		if (dir.isDirectory()) {
			String[] children = dir.list();
			// �ݹ�ɾ��Ŀ¼�е���Ŀ¼��
			for (int i = 0; i < children.length; i++) {
				boolean success = deleteDir(new File(dir, children[i]));
				if (!success) {
					return false;
				}
			}
		}
		// Ŀ¼��ʱΪ�գ�����ɾ��
		return dir.delete();
	}

	public void handle() {
		long start = System.currentTimeMillis();
		// System.out.println("===== ��ʼ�������� =====\n");
		jtaLog.setText("");
		LFWLog.writeLog(jtaLog, "=====  ��ʼ��������  =====");
		try {
			List<Map<String, String>> data = new ArrayList<Map<String, String>>();
			data = readSrc(dataFile);

			setProgressBar(mainWindow.getProgressBar(), 0);
			double total = data.size();
			total = total + total * LFWConst.COEFFICIENT;
			// System.out.println("total=" + total);

			int k = 0;
			for (Map<String, String> map : data) {
				// System.out.println("�����" + ++k + "������...");
				LFWLog.writeLog(jtaLog, "���ڴ����" + ++k + "������...");
				readwriteWord(modelFile, outPath, createFileType, map);
				int n = (int) ((double) k / total * 100);
				//System.out.println(n);
				setProgressBar(mainWindow.getProgressBar(), n);
			}

			if (createFileType == 'a') {

				File file = new File(outPath + "//" + LFWConst.FILE_TMP);

				String fn = outPath + "\\" + LFWStringUtil.getFileNameByFilePath(modelFile) + "_����"
						+ LFWConst.FILE_TYPE;
				// System.out.println("���ɽ���ļ���" + fn + "\n");
				LFWLog.writeLog(jtaLog, "���ɽ���ļ���" + fn + "");

				// �ϲ��ļ�
				mergeFiles(file, fn);
				// ɾ���ļ���
				deleteDir(file);
			}

		} catch (Exception e) {
			// System.out.println("���ݴ����쳣������ϵ����Ա��\n");
			LFWLog.writeLog(jtaLog, "���ݴ����쳣������ϵ����Ա��");
			e.printStackTrace();
		} finally {
			setProgressBar(mainWindow.getProgressBar(), 100);
			JButton btn = mainWindow.getBtnCreateFile();
			btn.setEnabled(true);
			btn.setText("�����ĵ�");
			LFWUtil.Alert("�ĵ�������ɣ�", JOptionPane.INFORMATION_MESSAGE);
			
		}

		// System.out.println("\n===== ���ݴ�����ϣ���ʱ��" + (System.currentTimeMillis()
		// - start) + " ���룩 =====");
		LFWLog.writeLog(jtaLog, "=====  ���ݴ�����ϣ���ʱ��" + (System.currentTimeMillis() - start) + " ���룩  =====");
	}

	private void setProgressBar(JProgressBar bar, int val) {
		int value = val; // bar.getValue() +
		if (value > bar.getMaximum()) {
			value = bar.getMaximum();
		}
		bar.setValue(value);
	}

	private String modelFile;
	private String dataFile;
	private String outPath;
	private char createFileType;
	private MainFramework mainWindow;

	private JTextArea jtaLog;

	public HandleFile(String modelFile, String dataFile, String outPath, char createFileType,
			MainFramework mainWindow) {
		super();
		this.modelFile = modelFile;
		this.dataFile = dataFile;
		this.outPath = outPath;
		this.createFileType = createFileType;
		this.mainWindow = mainWindow;

		jtaLog = mainWindow.getJtaLog();
	}

	@Override
	public void run() {
		handle();
	}
}
