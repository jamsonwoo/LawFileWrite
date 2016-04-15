package com.test;

import java.util.ArrayList;
import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TestDemo {
	public static void main(String[] args) {
		List list = new ArrayList();
		String file1 = "D:\\file1.doc";
		String file2 = "D:\\file2.doc";
		String file3 = "D:\\file3.doc";
		list.add(file1);
		list.add(file2);
		list.add(file3);
		uniteDoc(list, "d:\\file.doc");
	}

	public static void uniteDoc(List fileList, String savepaths) {
		if (fileList.size() == 0 || fileList == null) {
			return;
		}
		// ��word
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
							new int[3]).toDispatch();
			// ׷���ļ�
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
