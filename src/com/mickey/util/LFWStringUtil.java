package com.mickey.util;

public class LFWStringUtil {

	public static String getFilePath(String path) {
		if (path == null || "".equals(path)) {
			return null;
		}
		return path.substring(0, path.lastIndexOf("\\") + 1);
	}

	public static String getFileNameByFilePath(String path) {
		if (path == null || "".equals(path)) {
			return null;
		}
		// System.out.println("path=" + path);
		// System.out.println("path.lastIndexOf(\"\\\") =" +
		// path.lastIndexOf("\\") + 1);
		// System.out.println("path.lastIndexOf(\".\")=" +
		// path.lastIndexOf("."));

		return path.substring(path.lastIndexOf("\\") + 1, path.lastIndexOf("."));
	}

	public static void main(String[] args) {
		String path = "F:\\test\\结案报告表.doc";
		String s = path.substring(path.lastIndexOf("\\") + 1, path.lastIndexOf("."));
		System.out.println(s);
		// System.out.println(getFileNameByFilePath(path));
	}

}
