package com.mickey.core;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

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

	// private static int index = 1;

	// F:\test 结案报告表 名单.xls a

	private List<Map<String, String>> readSrc(String dataFile) throws Exception {
		// System.out.println("正在处理数据文件：" + dataFile + "\n");
		LFWLog.writeLog(jtaLog, "正在处理数据文件：" + dataFile + "");
		File file = new File(dataFile);

		List<Map<String, String>> data = new ArrayList<Map<String, String>>();
		data = ExcelOperate.getData(file);
		return data;

	}

	/**
	 * 实现对word读取和修改操作
	 * 
	 * @param filePath
	 *            word模板路径和名称
	 * @param map
	 *            待填充的数据，从数据库读取
	 */
	public void readwriteWord(String modelFile, String outPath,
			char createFileTypy, Map<String, String> map,
			Map<String, String> titleMap, int index) throws Exception {
		String strIndex = "000" + index;
		strIndex = strIndex.substring(strIndex.length() - 3);
		// 读取word模板
		FileInputStream in = null;
		ByteArrayOutputStream ostream = new ByteArrayOutputStream();
		FileOutputStream out = null;

		String dirName = LFWStringUtil.getFileNameByFilePath(modelFile);
		String outFileName = dirName + LFWConst.SPLIT;

		if (createFileTypy == 'a') {
			dirName = LFWConst.FILE_TMP;
			outFileName = LFWConst.FILE_TMP + LFWConst.SPLIT + strIndex;
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
			Iterator<Field> it = fields.getFields(FieldsDocumentPart.MAIN)
					.iterator();
			// while (it.hasNext()) {
			// System.out.println(it.next().getType());
			// }

			// 读取word文本内容
			Range range = hdt.getRange();
			/*
			 * TableIterator tableIt = new TableIterator(range); // 迭代文档中的表格
			 * while (tableIt.hasNext()) { Table tb = (Table) tableIt.next(); //
			 * 迭代行，默认从0开始 for (int i = 0; i < tb.numRows(); i++) { TableRow tr =
			 * tb.getRow(i); // 只读前8行，标题部分 if (i >= 8) break; // 迭代列，默认从0开始 for
			 * (int j = 0; j < tr.numCells(); j++) { TableCell td =
			 * tr.getCell(j);// 取得单元格 // 取得单元格的内容 for (int k = 0; k <
			 * td.numParagraphs(); k++) { Paragraph para = td.getParagraph(k);
			 * // String s = para.text(); // System.out.println(s); } } } }
			 */
			//
			// 替换文本内容
			for (Map.Entry<String, String> entry : map.entrySet()) {
				range.replaceText(entry.getKey(), entry.getValue());
			}
			// System.out.println(range.text());

			// 替换标签
			for (Map.Entry<String, String> entry : titleMap.entrySet()) {
				String keyTitle = entry.getKey();
				String keyData = keyTitle.replace("#", "$");
				// System.out.println(keyTitle + " : " + titleMap.get(keyTitle)
				// + "   " + keyData + " : " + map.get(keyData));
				if (map.get(keyData) != null && map.get(keyData) != "") {
					range.replaceText(entry.getKey(), entry.getValue());
				} else {
					range.replaceText(entry.getKey(), "");
				}
			}

			File folder = new File(outPath + "\\" + dirName);
			if (!(folder.exists() && folder.isDirectory())) {
				folder.mkdirs();
			}
			String fn = folder + "\\" + outFileName + map.get("${A}")
					+ LFWConst.FILE_TYPE;
			// System.out.println("生成结果文件：" + fn + "\n");
			LFWLog.writeLog(jtaLog, "生成结果文件：" + fn + "\n");
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

	private static void mergeFiles(File tmpFile, String createFile)
			throws IOException {
		// System.out.println(tmpFile.getPath());
		// System.out.println(createFile.getPath());
		File file = new File(createFile);
		if (file.exists()) {
			file.delete();
		}
		String[] children = tmpFile.list();
		// 递归删除目录中的子目录下
		for (int i = 0; i < children.length; i++) {
			// System.out.println(tmpFile+"\\"+children[i]);
			// list.add(tmpFile + "\\" + children[i]);
			children[i] = tmpFile + "\\" + children[i];
		}

		// MergeDoc.uniteDoc(list, createFile);
		MergeDoc mergeDoc = new MergeDoc();
		mergeDoc.uniteDoc(children, createFile);

	}

	private static boolean deleteDir(File dir) {
		if (dir.isDirectory()) {
			String[] children = dir.list();
			// 递归删除目录中的子目录下
			for (int i = 0; i < children.length; i++) {
				boolean success = deleteDir(new File(dir, children[i]));
				if (!success) {
					return false;
				}
			}
		}
		// 目录此时为空，可以删除
		return dir.delete();
	}

	public void handle() {
		long start = System.currentTimeMillis();
		// System.out.println("===== 开始处理数据 =====\n");
		jtaLog.setText("");
		LFWLog.writeLog(jtaLog, "=====  开始处理数据  =====");
		try {
			List<Map<String, String>> data = new ArrayList<Map<String, String>>();
			data = readSrc(dataFile);
			// System.out.println("=========");
			// System.out.println(data.get(0));
			// System.out.println("=========");

			// 设置表头
			Map<String, String> titleMap = new HashMap<String, String>();
			if (data != null && data.size() > 0) {
				Map<String, String> tmp = data.get(0);
				Set<String> keys = tmp.keySet();
				for (String key : keys) {
					titleMap.put(key.replace("$", "#"), tmp.get(key));
				}
				// System.out.println("title==========");
				// Set<String> k = titleMap.keySet();
				// for (String a : k) {
				// System.out.println(a + " : " + titleMap.get(a));
				// }
				//
				// System.out.println("title==========");
			}

			setProgressBar(mainWindow.getProgressBar(), 0);
			double total = data.size();
			total = total + total * LFWConst.COEFFICIENT;
			// System.out.println("total=" + total);

			int k = 0;
			for (Map<String, String> map : data) {
				if (k == 0) {
					k++;
					continue;
				} else {
					// System.out.println("处理第" + k + "个数据...");
					LFWLog.writeLog(jtaLog, "正在处理第" + k + "个数据...");
					readwriteWord(modelFile, outPath, createFileType, map,
							titleMap, k);
					k++;
					int n = (int) ((double) k / total * 100);
					// System.out.println(n);
					setProgressBar(mainWindow.getProgressBar(), n);
				}
			}

			if (createFileType == 'a') {

				File file = new File(outPath + "//" + LFWConst.FILE_TMP);

				String fn = outPath + "\\"
						+ LFWStringUtil.getFileNameByFilePath(modelFile)
						+ "_整合" + LFWConst.FILE_TYPE;
				// System.out.println("生成结果文件：" + fn + "\n");
				LFWLog.writeLog(jtaLog, "生成结果文件：" + fn + "");

				// 合并文件
				mergeFiles(file, fn);
				// 删除文件夹
				deleteDir(file);
			}

		} catch (Exception e) {
			// System.out.println("数据处理异常，请联系管理员。\n");
			LFWLog.writeLog(jtaLog, "数据处理异常，请联系管理员。");
			e.printStackTrace();
		} finally {
			setProgressBar(mainWindow.getProgressBar(), 100);
			JButton btn = mainWindow.getBtnCreateFile();
			btn.setEnabled(true);
			btn.setText("生成文档");
			LFWUtil.Alert("文档处理完成！", JOptionPane.INFORMATION_MESSAGE);

		}

		// System.out.println("\n===== 数据处理完毕（耗时：" + (System.currentTimeMillis()
		// - start) + " 毫秒） =====");
		LFWLog.writeLog(jtaLog, "=====  数据处理完毕（耗时："
				+ (System.currentTimeMillis() - start) + " 毫秒）  =====");
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

	public HandleFile(String modelFile, String dataFile, String outPath,
			char createFileType, MainFramework mainWindow) {
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
