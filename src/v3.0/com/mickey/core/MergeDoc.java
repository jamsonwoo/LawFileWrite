package com.mickey.core;

import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class MergeDoc {
	public static void main(String[] args) {
		/*
		 * List list = new ArrayList(); String file1 =
		 * "F:\\test\\tmp\\1_�С��.doc"; String file2 =
		 * "F:\\test\\tmp\\2_̷��ʢ.doc"; String file3 =
		 * "F:\\test\\tmp\\3_������.doc"; list.add(file1); list.add(file2);
		 * list.add(file3); uniteDoc(list, "f:\\test\\fileAll.doc");
		 */
	}

	private ActiveXComponent word;
	private Dispatch documents;
	private Dispatch doc;
	private Dispatch selection;

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
	public boolean uniteDoc(String[] filePath, String savepaths) {
		if (filePath == null || filePath.length == 0)
			return false;

		boolean successful = false;
		try {
			wordInit();

			// ���һ�����ĵ�
			doc = Dispatch.call(documents, "Add").toDispatch();

			selection = word.getProperty("Selection").toDispatch();

			for (int i = 0; i < filePath.length; i++) {
				// System.out.println("filePath = " + filePath[i]);
				Dispatch.call(selection, "insertFile", filePath[i]);
				if (i < filePath.length - 1) {
					Dispatch.call(selection, "TypeParagraph"); // ����
					Dispatch.call(selection, "InsertBreak"); // ��ҳ
				}
			}

			// �ļ����Ϊ
			Dispatch.call(doc, "SaveAs", savepaths);
			// Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method, new
			// Object[] { savepaths, new Variant(1) }, new int[3]);

			successful = true;

		} catch (Exception ex) {
			successful = false;
			ex.printStackTrace();
		} finally {
			wordDestory();
		}

		return successful;
	}

	public void uniteDoc(List<String> fileList, String savepaths) {
		if (fileList.size() == 0 || fileList == null) {
			return;
		} // ��word
		ActiveXComponent app = new ActiveXComponent("Word.Application");// ����word
		try {
			// ����word���ɼ�
			app.setProperty("Visible", new Variant(false));
			// ���documents����
			Object docs = app.getProperty("Documents").toDispatch();
			// �򿪵�һ���ļ�
			Object doc = Dispatch
					.invoke(
							(Dispatch) docs,
							"Open",
							Dispatch.Method,
							new Object[] { (String) fileList.get(0),
									new Variant(false), new Variant(true) },
							new int[3]).toDispatch(); // ׷���ļ�
			for (int i = 1; i < fileList.size(); i++) {
				Dispatch.invoke(app.getProperty("Selection").toDispatch(),
						"insertFile", Dispatch.Method, new Object[] {
								(String) fileList.get(i), "",
								new Variant(false), new Variant(false),
								new Variant(false) }, new int[3]);
			}
			// �����µ�word�ļ�
			Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method,
					new Object[] { savepaths, new Variant(1) }, new int[3]);
			Variant f = new Variant(false);
			Dispatch.call((Dispatch) doc, "Close", f);
		} catch (Exception e) {
			throw new RuntimeException("�ϲ�word�ļ�����.ԭ��:" + e);
		} finally {
			app.invoke("Quit", new Variant[] {});
		}
	}
}
