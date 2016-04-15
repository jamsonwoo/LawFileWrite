package com.mickey.core;

import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class MergeDoc {
	public static void main(String[] args) {
		/*
		 * List list = new ArrayList(); String file1 =
		 * "F:\\test\\tmp\\1_岑小敏.doc"; String file2 =
		 * "F:\\test\\tmp\\2_谭开盛.doc"; String file3 =
		 * "F:\\test\\tmp\\3_彭威豪.doc"; list.add(file1); list.add(file2);
		 * list.add(file3); uniteDoc(list, "f:\\test\\fileAll.doc");
		 */
	}

	private ActiveXComponent word;
	private Dispatch documents;
	private Dispatch doc;
	private Dispatch selection;

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
	public boolean uniteDoc(String[] filePath, String savepaths) {
		if (filePath == null || filePath.length == 0)
			return false;

		boolean successful = false;
		try {
			wordInit();

			// 添加一个新文挡
			doc = Dispatch.call(documents, "Add").toDispatch();

			selection = word.getProperty("Selection").toDispatch();

			for (int i = 0; i < filePath.length; i++) {
				// System.out.println("filePath = " + filePath[i]);
				Dispatch.call(selection, "insertFile", filePath[i]);
				if (i < filePath.length - 1) {
					Dispatch.call(selection, "TypeParagraph"); // 换行
					Dispatch.call(selection, "InsertBreak"); // 换页
				}
			}

			// 文件另存为
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
		} // 打开word
		ActiveXComponent app = new ActiveXComponent("Word.Application");// 启动word
		try {
			// 设置word不可见
			app.setProperty("Visible", new Variant(false));
			// 获得documents对象
			Object docs = app.getProperty("Documents").toDispatch();
			// 打开第一个文件
			Object doc = Dispatch
					.invoke(
							(Dispatch) docs,
							"Open",
							Dispatch.Method,
							new Object[] { (String) fileList.get(0),
									new Variant(false), new Variant(true) },
							new int[3]).toDispatch(); // 追加文件
			for (int i = 1; i < fileList.size(); i++) {
				Dispatch.invoke(app.getProperty("Selection").toDispatch(),
						"insertFile", Dispatch.Method, new Object[] {
								(String) fileList.get(i), "",
								new Variant(false), new Variant(false),
								new Variant(false) }, new int[3]);
			}
			// 保存新的word文件
			Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method,
					new Object[] { savepaths, new Variant(1) }, new int[3]);
			Variant f = new Variant(false);
			Dispatch.call((Dispatch) doc, "Close", f);
		} catch (Exception e) {
			throw new RuntimeException("合并word文件出错.原因:" + e);
		} finally {
			app.invoke("Quit", new Variant[] {});
		}
	}
}
