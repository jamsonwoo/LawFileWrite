package com.test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 将jacob.dll放入JDK的bin目录下 把jacob.jar放入项目的WEB-INF\lib目录下
 */
public class JacobHelper {
	private ActiveXComponent word;
	private Dispatch documents;
	private Dispatch doc;
	private Dispatch selection;

	private ActiveXComponent excel;
	private Dispatch workbooks;
	private Dispatch workbook;

	private void wordInit() {
		if (word == null) {
			// 创建一个word对象
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(false)); // word不可见
			word.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
		}

		if (documents == null) {
			// 获取文挡属性
			documents = word.getProperty("Documents").toDispatch();
		}
	}

	private void wordDestory() {
		try {
			if (doc != null) {
				// Dispatch.call(doc, "Close", new Variant(true));
				Dispatch.invoke(doc, "Close", Dispatch.Method,
						new Object[] { new Variant(true) }, new int[1]);
				doc = null;
			}
			if (word != null) {
				word.invoke("Quit", new Variant[] {});
				word.safeRelease();
				word = null;
				documents = null;
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	/**
	 * 合并Word文档
	 */
	public boolean wordUnite(String[] filePath, String outputPath) {
		if (filePath == null || filePath.length == 0)
			return false;

		boolean successful = false;
		try {
			wordInit();

			// 添加一个新文挡
			doc = Dispatch.call(documents, "Add").toDispatch();

			selection = word.getProperty("Selection").toDispatch();

			for (int i = 0; i < filePath.length; i++) {
				Dispatch.call(selection, "insertFile", filePath[i]);
				if (i < filePath.length - 1) {
					Dispatch.call(selection, "TypeParagraph"); // 换行
				}
			}

			// 文件另存为
			Dispatch.call(doc, "SaveAs", outputPath);

			successful = true;

		} catch (Exception ex) {
			successful = false;
			ex.printStackTrace();
		} finally {
			wordDestory();
		}

		return successful;
	}

	public static void main(String[] args) {
		JacobHelper helper = new JacobHelper();
		String file1 = "D:\\file1.doc";
		String file2 = "D:\\file2.doc";
		String file3 = "D:\\file3.doc";
		String[] list = { file1, file2, file3 };
		helper.wordUnite(list, "d:\\file.doc");
	}
}