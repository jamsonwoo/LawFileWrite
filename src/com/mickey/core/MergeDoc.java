package com.mickey.core;

import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class MergeDoc {
	public static void main(String[] args) {
		/*List list = new ArrayList();
		String file1 = "F:\\test\\tmp\\1_岑小敏.doc";
		String file2 = "F:\\test\\tmp\\2_谭开盛.doc";
		String file3 = "F:\\test\\tmp\\3_彭威豪.doc";
		list.add(file1);
		list.add(file2);
		list.add(file3);
		uniteDoc(list, "f:\\test\\fileAll.doc");*/
	}

	public static void uniteDoc(List<String> fileList, String savepaths) {
		if (fileList.size() == 0 || fileList == null) {
			return;
		}
		// 打开word
		ActiveXComponent app = new ActiveXComponent("Word.Application");// 启动word
		try {
			// 设置word不可见
			app.setProperty("Visible", new Variant(false));
			// 获得documents对象
			Object docs = app.getProperty("Documents").toDispatch();
			// 打开第一个文件
			Object doc = Dispatch.invoke((Dispatch) docs, "Open", Dispatch.Method,
					new Object[] { (String) fileList.get(0), new Variant(false), new Variant(true) }, new int[3])
					.toDispatch();
			// 追加文件
			for (int i = 1; i < fileList.size(); i++) {
				Dispatch.invoke(app.getProperty("Selection").toDispatch(), "insertFile", Dispatch.Method, new Object[] {
						(String) fileList.get(i), "", new Variant(false), new Variant(false), new Variant(false) },
						new int[3]);
			}
			// 保存新的word文件
			Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method, new Object[] { savepaths, new Variant(1) },
					new int[3]);
			Variant f = new Variant(false);
			Dispatch.call((Dispatch) doc, "Close", f);
		} catch (Exception e) {
			throw new RuntimeException("合并word文件出错.原因:" + e);
		} finally {
			app.invoke("Quit", new Variant[] {});
		}
	}
}
