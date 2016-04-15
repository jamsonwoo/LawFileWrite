package com.test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * ��jacob.dll����JDK��binĿ¼�� ��jacob.jar������Ŀ��WEB-INF\libĿ¼��
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
			// ����һ��word����
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(false)); // word���ɼ�
			word.setProperty("AutomationSecurity", new Variant(3)); // ���ú�
		}

		if (documents == null) {
			// ��ȡ�ĵ�����
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
	 * �ϲ�Word�ĵ�
	 */
	public boolean wordUnite(String[] filePath, String outputPath) {
		if (filePath == null || filePath.length == 0)
			return false;

		boolean successful = false;
		try {
			wordInit();

			// ���һ�����ĵ�
			doc = Dispatch.call(documents, "Add").toDispatch();

			selection = word.getProperty("Selection").toDispatch();

			for (int i = 0; i < filePath.length; i++) {
				Dispatch.call(selection, "insertFile", filePath[i]);
				if (i < filePath.length - 1) {
					Dispatch.call(selection, "TypeParagraph"); // ����
				}
			}

			// �ļ����Ϊ
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