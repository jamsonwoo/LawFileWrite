package com.test;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacob.com.ComThread;

public class StudyJacob {
	// word�ĵ�
	private Dispatch doc = null;
	// word���г������
	private ActiveXComponent word = null;
	// ����word�ĵ�����
	private Dispatch documents = null;
	// ѡ���ķ�Χ������
	private static Dispatch selection = null;
	// �����Ƿ񱣴����˳��ı�־
	private boolean saveOnExit = true;
	/**
	 * word�еĵ�ǰ�ĵ�
	 */
	private static Dispatch document = null;
	// private static Dispatch textFrame = null;
	//
	/**
	 * ���б��
	 */
	private Dispatch tables;
	/**
	 * ��ǰ���
	 */
	private Dispatch table;
	/**
	 * ��ǰ��Ԫ��
	 */
	private Dispatch cell;
	/**
	 * ��ǰ����е�������
	 */
	private Dispatch rows;
	/**
	 * ����е�ĳһ��
	 */
	private Dispatch row;
	/**
	 * ����е�������
	 */
	private Dispatch columns;
	/**
	 * ����е�ĳһ��
	 */
	private Dispatch column;
	/**
	 * ��wordʱͬʱҪ�򿪵��ĵ�����ָ��ʱ���½�һ���հ��ĵ�
	 */
	// private File openDoc;
	private static Dispatch shapes;
	private static Dispatch shape;
	private static Dispatch textRange;
	private static Dispatch textframe;
	private Dispatch range;
	private Dispatch paragraphs;
	private Dispatch paragraph;

	// constructor
	public StudyJacob() {
		if (word == null) {
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(true));
		}
		if (documents == null) {
			documents = word.getProperty("Documents").toDispatch();
		}
	}

	/**
	 * ����һ���µ�word�ĵ�
	 * 
	 */
	public void createNewDocument() {
		doc = Dispatch.call(documents, "Add").toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	/**
	 * ��һ���Ѵ��ڵ��ĵ�
	 * 
	 * @param docPath
	 */
	public void openDocument(String docPath) {
		if (this.doc != null) {
			this.closeDocument();
		}
		doc = Dispatch.call(documents, "Open", docPath).toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	/**
	 * �رյ�ǰword�ĵ�
	 * 
	 */
	public void closeDocument() {
		if (doc != null) {
			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", new Variant(saveOnExit));
			doc = null;
		}
	}

	/**
	 * �ر�ȫ��Ӧ��
	 * 
	 */
	public void close() {
		closeDocument();
		if (word != null) {
			Dispatch.call(word, "Quit");
			word = null;
		}
		selection = null;
		documents = null;
	}

	/**
	 * �Ѳ�����ƶ����ļ���λ��
	 * 
	 */
	public void moveStart() {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		Dispatch.call(selection, "HomeKey", new Variant(6));
	}

	/**
	 * �ڵ�ǰ���������ַ���
	 * 
	 * @param newText
	 *            Ҫ��������ַ���
	 */
	public void insertText(String newText) {
		Dispatch.put(selection, "Text", newText);
	}

	/**
	 * �ڵ�ǰ��������ͼƬ
	 * 
	 * @param imagePath
	 *            ͼƬ·��
	 */
	public void insertImage(String imagePath) {
		Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(),
				"AddPicture", imagePath);
	}

	/**
	 * ��ѡ�������ݻ��߲���������ƶ�
	 * 
	 * @param pos
	 *            �ƶ��ľ���
	 */
	public void moveDown(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveDown");
	}

	/**
	 * ��ѡ�������ݻ����������ƶ�
	 * 
	 * @param pos
	 *            �ƶ��ľ���
	 */
	public void moveUp(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveUp");
	}

	/**
	 * ��ѡ�������ݻ��߲���������ƶ�
	 * 
	 * @param pos
	 *            �ƶ��ľ���
	 */
	public void moveLeft(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++) {
			Dispatch.call(selection, "MoveLeft");
		}
	}

	/**
	 * ��ѡ�������ݻ��߲���������ƶ�
	 * 
	 * @param pos
	 *            �ƶ��ľ���
	 */
	public void moveRight(int pos) {
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveRight");
	}

	/**
	 * �ļ���������Ϊ
	 * 
	 * @param savePath
	 *            һ��Ҫ�ǵü�����չ�� .doc ��������Ϊ·��
	 */
	public void save(String savePath) {
		Dispatch.call(
				(Dispatch) Dispatch.call(word, "WordBasic").getDispatch(),
				"FileSaveAs", savePath);
	}

	/**
	 * �ӵ�tIndex��Table��ȡ��ֵ��row�У���col�е�ֵ
	 * 
	 * @param tableIndex
	 *            �ĵ��еĵ�tIndex��Table����tIndexΪ����ȡ
	 * @param cellRowIdx
	 *            cell��Table��row��
	 * @param cellColIdx
	 *            cell��Talbe��col��
	 * @return cell��Ԫֵ
	 * @throws Exception
	 */
	public String getCellString(int tableIndex, int cellRowIdx, int cellColIdx)
			throws Exception {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫȡ���ݵı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),
				new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		return Dispatch.get(selection, "Text").toString();
	}

	/**
	 * �ӵ�tableIndex��Table��ȡ��ֵ��cellRowIdx�У���cellColIdx�е�ֵ
	 * 
	 * @param tIndex
	 *            �ĵ��еĵ�tIndex��Table����tIndexΪ����ȡ
	 * @param cellRowIdx
	 *            cell��Table��row��
	 * @param cellColIdx
	 *            cell��Talbe��col��
	 * @return cell��Ԫֵ
	 * @throws Exception
	 */
	public void getCellValue(int tableIndex, int cellRowIdx, int cellColIdx)
			throws Exception {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫȡ���ݵı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),
				new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		Dispatch.call(selection, "Copy");
	}

	/**
	 * �ڵ�ǰ��괦��ճ��
	 */
	public void paste() {
		Dispatch.call(selection, "Paste");
	}

	/**
	 * �ڵ�ǰ��괦���ͼƬ
	 * 
	 * @param imgPath
	 *            ͼƬ�ĵ�ַ
	 */
	public void addImage(String imgPath) {
		if (imgPath != "" && imgPath != null) {
			Dispatch image = Dispatch.get(selection, "InLineShapes")
					.toDispatch();
			Dispatch.call(image, "AddPicture", imgPath);
		}
	}

	/**
	 * ��ָ���ĵ�Ԫ������д����
	 * 
	 * @param tableIndex
	 *            �ĵ��еĵ�tIndex��Table����tIndexΪ����ȡ
	 * @param cellRowIdx
	 *            cell��Table��row��
	 * @param cellColIdx
	 *            cell��Talbe��col��
	 * @param txt
	 *            ��д������
	 */
	public void putTxtToCell(int tableIndex, int cellRowIdx, int cellColIdx,
			String txt) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),
				new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		Dispatch.put(selection, "Text", txt);
	}

	/**
	 * 
	 * �õ���ǰ�ĵ���tables����
	 */
	public Dispatch getTables() throws Exception {
		if (this.doc == null) {
			throw new Exception("there is not a document can't be operate!!!");
		}
		return Dispatch.get(doc, "Tables").toDispatch();
	}

	/**
	 * 
	 * �õ���ǰ�ĵ��ı����
	 * 
	 * @param Dispatch
	 */
	public int getTablesCount(Dispatch tables) throws Exception {
		int count = 0;
		try {
			this.getTables();
		} catch (Exception e) {
			throw new Exception("there is not any table!!");
		}
		count = Dispatch.get(tables, "Count").toInt();
		return count;
	}

	/**
	 * �ڵ�ǰ�ĵ�ָ����λ�ÿ������
	 * 
	 * @param pos
	 *            ��ǰ�ĵ�ָ����λ��
	 * @param tableIndex
	 *            �������ı����word�ĵ���������λ��
	 */
	public void copyTable(String pos, int tableIndex) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		Dispatch range = Dispatch.get(table, "Range").toDispatch();
		Dispatch.call(range, "Copy");
		Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
		Dispatch.call(textRange, "Paste");
	}

	/**
	 * �ڵ�ǰ�ĵ�ָ����λ�ÿ���������һ���ĵ��еı��
	 * 
	 * @param anotherDocPath
	 *            ��һ���ĵ��Ĵ���·��
	 * @param tableIndex
	 *            �������ı������һ���ĵ��е�λ��
	 * @param pos
	 *            ��ǰ�ĵ�ָ����λ��
	 */
	public void copyTableFromAnotherDoc(String anotherDocPath, int tableIndex,
			String pos) {
		Dispatch doc2 = null;
		try {
			doc2 = Dispatch.call(documents, "Open", anotherDocPath)
					.toDispatch();
			// ���б��
			Dispatch tables = Dispatch.get(doc2, "Tables").toDispatch();
			// Ҫ���ı��
			Dispatch table = Dispatch.call(tables, "Item",
					new Variant(tableIndex)).toDispatch();
			Dispatch range = Dispatch.get(table, "Range").toDispatch();
			Dispatch.call(range, "Copy");
			Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
			moveDown(1);
			Dispatch.call(textRange, "Paste");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (doc2 != null) {
				Dispatch.call(doc2, "Close", new Variant(saveOnExit));
				doc2 = null;
			}
		}
	}

	/**
	 * �ڵ�ǰ�ĵ���������������
	 * 
	 * @param pos
	 */
	public void pasteExcelSheet(String pos) {
		moveStart();
		Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
		Dispatch.call(textRange, "Paste");
	}

	/**
	 * ѡ��dispatch����
	 * 
	 * @param dispatch
	 *            �����䣬��ǲ��
	 */
	private void select(Dispatch dispatch) {
		Dispatch.call(dispatch, "Select");
	}

	/**
	 * �ڵ�ǰ�ĵ�ָ����λ�ÿ���������һ���ĵ��е�ͼƬ
	 * 
	 * @param anotherDocPath
	 *            ��һ���ĵ��Ĵ���·��
	 * @param shapeIndex
	 *            ��������ͼƬ����һ���ĵ��е�λ��
	 * @param pos
	 *            ��ǰ�ĵ�ָ����λ��
	 * @throws com.jacob.com.ComFailException
	 *             Invoke of: Item Source: Microsoft Word ��shapeIndex������
	 */
	public void copyImageFromAnotherDoc(String anotherDocPath, int shapeIndex,
			String pos) {
		Dispatch doc2 = null;
		try {
			doc2 = Dispatch.call(documents, "Open", anotherDocPath)
					.toDispatch();
			Dispatch shapes = Dispatch.get(doc2, "InLineShapes").toDispatch();
			Dispatch shape = Dispatch.call(shapes, "Item",
					new Variant(shapeIndex)).toDispatch();
			Dispatch imageRange = Dispatch.get(shape, "Range").toDispatch();
			Dispatch.call(imageRange, "Copy");
			Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
			moveDown(4);
			Dispatch.call(textRange, "Paste");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (doc2 != null) {
				Dispatch.call(doc2, "Close", new Variant(saveOnExit));
				doc2 = null;
			}
		}
	}

	/**
	 * ��ӡ��ǰword�ĵ�
	 * 
	 * @throws Exception
	 *             com.jacob.com.ComFailException: Invoke of: PrintOut Source:
	 *             Microsoft Word ���޴�ӡ��
	 */
	public void printFile() {
		if (doc != null) {
			Dispatch.call(doc, "PrintOut");
		}
	}

	/**
	 * ��ӡ�ı�����ѡ�����ı�����<br>
	 * ���棺ʹ����Home, End��ȡ��ѡ��
	 * 
	 * @param s
	 */
	public void println(String s) {
		write(s);
		goToEnd();
		cmd("TypeParagraph");
	}

	/**
	 * ִ��ĳ����ָ��
	 * 
	 * @param cmd
	 */
	private void cmd(String cmd) {
		Dispatch.call(selection, cmd);
	}

	/**
	 * ����Ctrl + End��
	 */
	public void goToEnd() {
		Dispatch.call(selection, "EndKey", "6");
	}

	/**
	 * ��ѡ���ٴ�ӡ�ı�<br>
	 * ���棺ʹ����Home, End��ȡ��ѡ��
	 */
	public void print(String s) {
		goToEnd();
		write(s);
	}

	/**
	 * ����ѡ, ֱ������ı�
	 * 
	 * @param s
	 */
	public void write(String s) {
		Dispatch.put(selection, "Text", s);
	}

	/**
	 * ��ѡ�����ı�����<br>
	 * ���棺ʹ����Home, End��ȡ��ѡ��
	 */
	public void println() {
		home();
		end();
		cmd("TypeParagraph");
	}

	/**
	 * ����Home��
	 */
	public void home() {
		Dispatch.call(selection, "HomeKey", "5");
	}

	/**
	 * ����End��
	 */
	public void end() {
		Dispatch.call(selection, "EndKey", "5");
	}

	/**
	 * ����Ctrl + Home��
	 */
	public void goToBegin() {
		Dispatch.call(selection, "HomeKey", "6");
	}

	/**
	 * ����ָ�����ָ���е��п�
	 * 
	 * @param tableIndex
	 * @param columnWidth
	 * @param columnIndex
	 */
	public void setColumnWidth(int tableIndex, float columnWidth,
			int columnIndex) {
		this.getTable(tableIndex);
		this.setColumnWidth(columnWidth, columnIndex);
	}

	/**
	 * ���ұ�
	 * 
	 * @param tableIndex
	 * @return
	 */
	public Dispatch getTable(int tableIndex) {
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		return table;
	}

	public Dispatch getShapes() throws Exception {
		return Dispatch.get(doc, "Shapes").toDispatch();
	}

	public int getShapesCount() throws Exception {
		int count = 0;
		count = Dispatch.get(shapes, "Count").toInt();
		return count;
	}

	public Dispatch getShape(int tIndex) throws Exception {
		return Dispatch.call(shapes, "Item", new Variant(tIndex)).toDispatch();
		// return Dispatch.invoke(shapes,"item",Dispatch.Method,new Object[]{
		// new Integer(tIndex)},new int[1]).toDispatch();
	}

	public Dispatch getTextFrame() throws Exception {
		return Dispatch.get(shape, "TextFrame").toDispatch();
	}

	public Dispatch getTextRange() throws Exception {
		return Dispatch.get(textframe, "TextRange").toDispatch();
	}

	/**
	 * ���õ�ǰ���ָ���е��п�
	 * 
	 * @param columnWidth
	 * @param columnIndex
	 * @throws �����������ı����ʹ��
	 */
	public void setColumnWidth(float columnWidth, int columnIndex) {
		if (columnWidth < 11) {
			columnWidth = 120;
		}
		if (columns == null || column == null) {
			this.getColumns();
			this.getColumn(columnIndex);
		}
		Dispatch.put(column, "Width", new Variant(columnWidth));
	}

	/**
	 * ����ָ�����ָ���еı���ɫ
	 * 
	 * @param tableIndex
	 * @param columnIndex
	 * @param color
	 *            ȡֵ��Χ 0 < color < 17 Ĭ�ϣ�16 ǳ��ɫ 1����ɫ 2����ɫ 3��ǳ�� ...............
	 */
	public void setColumnBgColor(int tableIndex, int columnIndex, int color) {
		this.getTable(tableIndex);
		this.setColumnBgColor(columnIndex, color);
	}

	/**
	 * ���õ�ǰ���ָ���еı���ɫ
	 * 
	 * @param columnIndex
	 * @param color
	 *            ȡֵ��Χ 0 < color < 17 Ĭ�ϣ�16 ǳ��ɫ 1����ɫ 2����ɫ 3��ǳ�� ...............
	 */
	public void setColumnBgColor(int columnIndex, int color) {
		this.getColumn(columnIndex);
		Dispatch shading = Dispatch.get(column, "Shading").toDispatch();
		if (color > 16 || color < 1)
			color = 16;
		Dispatch
				.put(shading, "BackgroundPatternColorIndex", new Variant(color));
	}

	/**
	 * ��ʼ�� com �߳�
	 */
	public void initCom() {
		ComThread.InitSTA();
	}

	/**
	 * �ͷ� com �߳���Դ com ���̻߳��ղ��� java �������ջ��ƻ���
	 */
	public void releaseCom() {
		ComThread.Release();
	}

	/**
	 * ���õ�ǰ���ָ���еı���ɫ
	 * 
	 * @param rowIndex
	 * @param color
	 *            ȡֵ��Χ 0 < color < 17 Ĭ�ϣ�16 ǳ��ɫ 1����ɫ 2����ɫ 3��ǳ�� ...............
	 */
	public void setRowBgColor(int rowIndex, int color) {
		this.getRow(rowIndex);
		Dispatch shading = Dispatch.get(row, "Shading").toDispatch();
		if (color > 16 || color < 1)
			color = 16;
		Dispatch
				.put(shading, "BackgroundPatternColorIndex", new Variant(color));
	}

	/**
	 * ����ָ������ָ���еı���ɫ
	 * 
	 * @param tableIndex
	 * @param rowIndex
	 * @param color
	 *            ȡֵ��Χ 0 < color < 17 Ĭ�ϣ�16 ǳ��ɫ 1����ɫ 2����ɫ 3��ǳ�� ...............
	 */
	public void setRowBgColor(int tableIndex, int rowIndex, int color) {
		this.getTable(tableIndex);
		this.setRowBgColor(rowIndex, color);
	}

	/**
	 * ���õ�ǰѡ�����ݵ�����
	 * 
	 * @param isBold
	 *            �Ƿ�Ϊ����
	 * @param isItalic
	 *            �Ƿ�Ϊб��
	 * @param isUnderLine
	 *            �Ƿ���»���
	 * @param color
	 *            rgb ������ɫ ���磺��ɫ 255,0,0
	 * @param size
	 *            �����С 12:С�� 16:����
	 * @param name
	 *            �������� ���磺���壬�����壬���壬����
	 */
	public void setFont(boolean isBold, boolean isItalic, boolean isUnderLine,
			String color, String size, String name) {
		Dispatch font = Dispatch.get(selection, "Font").toDispatch();
		Dispatch.put(font, "Name", new Variant(name));
		Dispatch.put(font, "Bold", new Variant(isBold));
		Dispatch.put(font, "Italic", new Variant(isItalic));
		Dispatch.put(font, "Underline", new Variant(isUnderLine));
		Dispatch.put(font, "Color", color);
		Dispatch.put(font, "Size", size);
	}

	/**
	 * �ָ�Ĭ������ ���Ӵ֣�����б��û�»��ߣ���ɫ��С�ĺ��֣�����
	 */
	public void clearFont() {
		this.setFont(false, false, false, "0,0,0", "12", "����");
	}

	/**
	 * �Ե�ǰ������и�ʽ��
	 * 
	 * @param align
	 *            �������з�ʽ Ĭ�ϣ����� 0:���� 1:���� 2:���� 3:���˶��� 4:��ɢ����
	 * @param lineSpace
	 *            �����м�� Ĭ�ϣ�1.0 0��1.0 1��1.5 2��2.0 3����Сֵ 4���̶�ֵ
	 */
	public void setParaFormat(int align, int lineSpace) {
		if (align < 0 || align > 4) {
			align = 0;
		}
		if (lineSpace < 0 || lineSpace > 4) {
			lineSpace = 0;
		}
		Dispatch alignment = Dispatch.get(selection, "ParagraphFormat")
				.toDispatch();
		Dispatch.put(alignment, "Alignment", align);
		Dispatch.put(alignment, "LineSpacingRule", new Variant(lineSpace));
	}

	/**
	 * ��ԭ����Ĭ�ϵĸ�ʽ �����,�м�ࣺ1.0
	 */
	public void clearParaFormat() {
		this.setParaFormat(0, 0);
	}

	/**
	 * �������
	 * 
	 * @param pos
	 *            λ��
	 * @param cols
	 *            ����
	 * @param rows
	 *            ����
	 */
	public void createTable(String pos, int numCols, int numRows) {
		if (find(pos)) {
			Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
			Dispatch range = Dispatch.get(selection, "Range").toDispatch();
			Dispatch newTable = Dispatch.call(tables, "Add", range,
					new Variant(numRows), new Variant(numCols)).toDispatch();
			Dispatch.call(selection, "MoveRight");
		}
	}

	/**
	 * ��ָ����ǰ��������
	 * 
	 * @param tableIndex
	 *            word�ļ��еĵ�N�ű�(��1��ʼ)
	 * @param rowIndex
	 *            ָ���е����(��1��ʼ)
	 */
	public void addTableRow(int tableIndex, int rowIndex) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.call(rows, "Item", new Variant(rowIndex))
				.toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * �ڵ�1��ǰ����һ��
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addFirstTableRow(int tableIndex) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.get(rows, "First").toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * �����1��ǰ����һ��
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addLastTableRow(int tableIndex) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.get(rows, "Last").toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * ����һ��
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addRow(int tableIndex) {
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch.call(rows, "Add");
	}

	/**
	 * ����һ��
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addCol(int tableIndex) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch.call(cols, "Add").toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * ��ָ����ǰ�����ӱ�����
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 * @param colIndex
	 *            �ƶ��е���� (��1��ʼ)
	 */
	public void addTableCol(int tableIndex, int colIndex) {
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		System.out.println(Dispatch.get(cols, "Count"));
		Dispatch col = Dispatch.call(cols, "Item", new Variant(colIndex))
				.toDispatch();
		// Dispatch col = Dispatch.get(cols, "First").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * �ڵ�1��ǰ����һ��
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addFirstTableCol(int tableIndex) {
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch col = Dispatch.get(cols, "First").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * �����һ��ǰ����һ��
	 * 
	 * @param tableIndex
	 *            word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addLastTableCol(int tableIndex) {
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))
				.toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch col = Dispatch.get(cols, "Last").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * ��ѡ�����ݻ����㿪ʼ�����ı�
	 * 
	 * @param toFindText
	 *            Ҫ���ҵ��ı�
	 * @return boolean true-���ҵ���ѡ�и��ı���false-δ���ҵ��ı�
	 */
	public boolean find(String toFindText) {
		if (toFindText == null || toFindText.equals(""))
			return false;
		// ��selection����λ�ÿ�ʼ��ѯ
		Dispatch find = word.call(selection, "Find").toDispatch();
		// ����Ҫ���ҵ�����
		Dispatch.put(find, "Text", toFindText);
		// ��ǰ����
		Dispatch.put(find, "Forward", "True");
		// ���ø�ʽ
		Dispatch.put(find, "Format", "True");
		// ��Сдƥ��
		Dispatch.put(find, "MatchCase", "True");
		// ȫ��ƥ��
		Dispatch.put(find, "MatchWholeWord", "True");
		// ���Ҳ�ѡ��
		return Dispatch.call(find, "Execute").getBoolean();
	}

	/**
	 * ��ѡ��ѡ�������趨Ϊ�滻�ı�
	 * 
	 * @param toFindText
	 *            �����ַ���
	 * @param newText
	 *            Ҫ�滻������
	 * @return
	 */
	public boolean replaceText(String toFindText, String newText) {
		if (!find(toFindText))
			return false;
		Dispatch.put(selection, "Text", newText);
		return true;
	}

	/**
	 * ȫ���滻�ı�
	 * 
	 * @param toFindText
	 *            �����ַ���
	 * @param newText
	 *            Ҫ�滻������
	 */
	public void replaceAllText(String toFindText, String newText) {
		while (find(toFindText)) {
			Dispatch.put(selection, "Text", newText);
			Dispatch.call(selection, "MoveRight");
		}
	}

	/**
	 * 
	 * @param toFindText
	 *            Ҫ���ҵ��ַ���
	 * @param imagePath
	 *            ͼƬ·��
	 * @return �˺������ַ����滻��ͼƬ
	 */
	public boolean replaceImage(String toFindText, String imagePath) {
		if (!find(toFindText))
			return false;
		Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(),
				"AddPicture", imagePath);
		return true;
	}

	/**
	 * ȫ���滻ͼƬ
	 * 
	 * @param toFindText
	 *            �����ַ���
	 * @param imagePath
	 *            ͼƬ·��
	 */
	public void replaceAllImage(String toFindText, String imagePath) {
		while (find(toFindText)) {
			Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(),
					"AddPicture", imagePath);
			Dispatch.call(selection, "MoveRight");
		}
	}

	// ////////////////////////////////////////////////////////
	// //////////////////////////////////////////////////////////
	// /////////////////////////////////////////////////////
	// ///////////////////////////////////////////////////
	// //////////////////////////////////////////////////
	/**
	 * ���õ�ǰ����ߵĴ�ϸ w��Χ��1<w<13 ������Χ��Ϊ��w=6
	 * 
	 * @param w
	 */
	/*
	 * private void setTableBorderWidth(int w) { if (w > 13 || w < 2) { w = 6; }
	 * Dispatch borders = Dispatch.get(table, "Borders").toDispatch(); Dispatch
	 * border = null;
	 * 
	 * /** ���ñ���ߵĴ�ϸ 1���������ϱ�һ���� 2�����������һ���� 3�����±�һ���� 4�����ұ�һ���� 5�������ϱ����±�֮������к���
	 * 6������������ұ�֮����������� 7�������Ͻǵ����½ǵ�б�� 8�������½ǵ����Ͻǵ�б��
	 */
	/*
	 * for (int i = 1; i < 7; i++) { border = Dispatch.call(borders, "Item", new
	 * Variant(i)) .toDispatch(); Dispatch.put(border, "LineWidth", new
	 * Variant(w)); Dispatch.put(border, "Visible", new Variant(true)); } }
	 * 
	 * 
	 * 
	 * /** ���Ʊ�����һ�е����а�
	 * 
	 * @param tableIndex
	 */
	/*
	 * public void copyLastRow(int tableIndex) {
	 * getRow(getRowsCount(tableIndex)); Dispatch.call(row, "select");
	 * Dispatch.call(selection, "Copy"); }
	 * 
	 * /** ���Ʊ�����һ�в�ճ������һ�У��������е����ݣ�
	 * 
	 * @param tableIndex ������� @param times ճ���Ĵ���
	 */
	/*
	 * public void duplicateLastRow(int tableIndex, int times) {
	 * this.copyLastRow(tableIndex); for (int i = 0; i < times; i++) {
	 * Dispatch.call(selection, "Paste"); } }
	 * 
	 * /** ���ҵ�ǰ�б���������е�ĳһ��
	 * 
	 * @param rowIndex @return
	 */
	public Dispatch getRow(int rowIndex) {
		if (rows == null)
			this.getRows();
		row = Dispatch.invoke(rows, "item", Dispatch.Method,
				new Object[] { new Integer(rowIndex) }, new int[1])
				.toDispatch();
		return row;
	}

	public int getRowsCount() {
		if (rows == null)
			this.getRows();
		return Dispatch.get(rows, "Count").getInt();
	}

	/**
	 * �õ���ǰ�������е���
	 * 
	 * @return
	 */
	// ��Ҫ�ҵ�Dispatch����,�����Variant(1)����һ��Ҫ���ɱ���
	public Dispatch getColumns() {
		// this.getTables();
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		Dispatch table = Dispatch.call(tables, "Item", new Variant(1))
				.toDispatch();
		return this.columns = Dispatch.get(table, "Columns").toDispatch();
	}

	/**
	 * �õ���ǰ����ĳһ��
	 * 
	 * @param index
	 *            ������
	 * @return
	 */
	public Dispatch getColumn(int columnIndex) {
		if (columns == null)
			this.getColumns();
		return this.column = Dispatch.call(columns, "Item",
				new Variant(columnIndex)).toDispatch();
	}

	/**
	 * �õ���ǰ��������
	 * 
	 * @return
	 */
	public int getColumnsCount() {
		this.getColumns();
		return Dispatch.get(columns, "Count").toInt();
	}

	/**
	 * �õ�ָ����������
	 * 
	 * @param tableIndex
	 * @return
	 */
	public int getColumnsCount(int tableIndex) {
		this.getTable(tableIndex);
		return this.getColumnsCount();
	}

	/**
	 * �õ��������
	 * 
	 * @param tableIndex
	 * @return
	 */
	public int getRowsCount(int tableIndex) {
		this.getTable(tableIndex);
		return this.getRowsCount();
	}

	/**
	 * ���õ�ǰ���������е��и�
	 * 
	 * @param rowHeight
	 */
	public void setRowHeight(float rowHeight) {
		if (rowHeight > 0) {
			if (rows == null)
				this.getRows();
			Dispatch.put(rows, "Height", new Variant(rowHeight));
		}
	}

	/**
	 * ����ָ�����������е��и�
	 * 
	 * @param tableIndex
	 * @param rowHeight
	 */
	public void setRowHeight(int tableIndex, float rowHeight) {
		this.getRows(tableIndex);
		this.setRowHeight(rowHeight);
	}

	/**
	 * ���õ�ǰ���ָ���е��и�
	 * 
	 * @param rowHeight
	 * @param rowIndex
	 */
	public void setRowHeight(float rowHeight, int rowIndex) {
		if (rowHeight > 0) {
			if (rows == null || row == null) {
				this.getRows();
				this.getRow(rowIndex);
			}
			Dispatch.put(row, "Height", new Variant(rowHeight));
		}
	}

	/**
	 * ����ָ������ָ���е��и�
	 * 
	 * @param tableIndex
	 * @param rowHeight
	 * @param rowIndex
	 */
	public void setRowHeight(int tableIndex, float rowHeight, int rowIndex) {
		this.getTable(tableIndex);
		this.setRowHeight(rowHeight, rowIndex);
	}

	/**
	 * ���õ�ǰ���������е��п�
	 * 
	 * @param columnWidth
	 *            �п� ȡֵ��Χ��10<columnWidth Ĭ��ֵ��120
	 */
	public void setColumnWidth(float columnWidth) {
		if (columnWidth < 11) {
			columnWidth = 120;
		}
		if (columns == null)
			this.getColumns();
		Dispatch.put(columns, "Width", new Variant(columnWidth));
	}

	/**
	 * ����ָ�����������е��п�
	 * 
	 * @param tableIndex
	 * @param columnWidth
	 */
	public void setColumnWidth(int tableIndex, float columnWidth) {
		this.getColumns(tableIndex);
		this.setColumnWidth(columnWidth);
	}

	/**
	 * �õ�ָ�����Ķ�����
	 * 
	 * @param tableIndex
	 * @return
	 */
	public Dispatch getColumns(int tableIndex) {
		getTable(tableIndex);
		return this.getColumns();
	}

	/**
	 * ���Ʊ��ĳһ��
	 * 
	 * @param tableIndex
	 * @param rowIndex
	 */
	public void copyRow(int tableIndex, int rowIndex) {
		getTable(tableIndex);
		getRows();
		row = getRow(rowIndex);
		Dispatch.call(row, "Select");
		Dispatch.call(selection, "Copy");
	}

	/**
	 * ���ұ��ȫ����
	 * 
	 * @param tableIndex
	 * @return
	 */
	public Dispatch getRows(int tableIndex) {
		getTable(tableIndex);
		return this.getRows();
	}

	/**
	 * ���ҵ�ǰ���ȫ����
	 * 
	 * @return
	 * 
	 * 
	 */
	// ��Ҫ�ҵ�Dispatch����,�����Variant(1)����һ��Ҫ���ɱ���
	public Dispatch getRows() {
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		Dispatch table = Dispatch.call(tables, "Item", new Variant(2))
				.toDispatch();
		rows = Dispatch.get(table, "rows").toDispatch();
		return rows;
	}

	/**
	 * ����ָ�����ĵ�Ԫ��
	 * 
	 * @param tableIndex
	 * @param cellRow
	 * @param cellColumn
	 * @return
	 */
	public Dispatch getCell(int tableIndex, int cellRow, int cellColumn) {
		getTable(tableIndex);
		return getCell(cellRow, cellColumn);
	}

	/**
	 * ���ҵ�ǰ���ڱ��ĳ��Ԫ��
	 * 
	 * @param cellRow
	 * @param cellColumn
	 * @return
	 * @throws Dispatch
	 *             object expected
	 */
	public Dispatch getCell(int cellRow, int cellColumn) {
		return cell = Dispatch.call(table, "Cell", new Variant(cellRow),
				new Variant(cellColumn)).toDispatch();
	}

	/**
	 * �����ĵ����˳�����
	 * 
	 * @param fileName
	 *            ������ļ�����
	 * @param isSave
	 *            �Ƿ񱣴��޸�
	 * @throws Exception
	 */
	/*
	 * public void saveDocAndExit(File fileName, boolean isSave) throws
	 * Exception { if (fileName != null) { if (!fileName.exists()) {
	 * fileName.createNewFile(); } Dispatch wordBasic = (Dispatch)
	 * Dispatch.call(word, "WordBasic").getDispatch();
	 * Dispatch.invoke(wordBasic, "FileSaveAs", Dispatch.Method, new Object[] {
	 * fileName.getPath(), new Variant(true), new Variant(false) }, new int[1]);
	 * }
	 * 
	 * Dispatch.call(document, "Close", new Variant(isSave));
	 * Dispatch.call(word, "Quit");
	 * 
	 * word = null; documents = null; document = null; selection = null; }
	 */
	/**
	 * �ϲ���ǰ���ָ���ĵ�Ԫ�� �����Ҫһ�κϲ�������Ԫ��ֻ��Ҫָ����һ����Ԫ������һ����Ԫ��
	 * 
	 * @param fstCellRowIndex
	 *            ��һ����Ԫ���������
	 * @param fstCellColIndex
	 *            ��һ����Ԫ���������
	 * @param secCellRowIndex
	 *            �ڶ�����Ԫ���������
	 * @param secCellColIndex
	 *            �ڶ�����Ԫ���������
	 */
	public void mergeCell(int fstCellRowIndex, int fstCellColIndex,
			int secCellRowIndex, int secCellColIndex) {
		Dispatch fstCell = Dispatch.call(table, "Cell",
				new Variant(fstCellRowIndex), new Variant(fstCellColIndex))
				.toDispatch();
		Dispatch secCell = Dispatch.call(table, "Cell",
				new Variant(secCellRowIndex), new Variant(secCellColIndex))
				.toDispatch();
		Dispatch.call(fstCell, "Merge", secCell);
	}

	/**
	 * �ϲ���ǰ���ָ������
	 * 
	 * @param columnIndex
	 *            ������
	 */
	public void mergeColumn(int columnIndex) {
		this.getColumn(columnIndex);
		Dispatch cells = Dispatch.get(column, "Cells").toDispatch();
		Dispatch.call(cells, "Merge");
	}

	/**
	 * �ϲ���ǰ����ָ������
	 * 
	 * @param rowIndex
	 */
	public void mergeRow(int rowIndex) {
		this.getRow(rowIndex);
		Dispatch cells = Dispatch.get(row, "Cells").toDispatch();
		Dispatch.call(cells, "Merge");
	}

	/**
	 * �ϲ�ָ������ָ������
	 * 
	 * @param tableIndex
	 * @param rowIndex
	 *            ������
	 */
	public void mergeRow(int tableIndex, int rowIndex) {
		this.getTable(tableIndex);
		this.mergeRow(rowIndex);
	}

	/**
	 * �ϲ�ָ������ָ������
	 * 
	 * @param tableIndex
	 * @param columnIndex
	 */
	public void mergeColumn(int tableIndex, int columnIndex) {
		this.getTable(tableIndex);
		this.mergeColumn(columnIndex);
	}

	/**
	 * �ϲ�ָ������ָ���ĵ�Ԫ��
	 * 
	 * @param tableIndex
	 * @param fstCellRowIndex
	 * @param fstCellColIndex
	 * @param secCellRowIndex
	 * @param secCellColIndex
	 */
	public void mergeCell(int tableIndex, int fstCellRowIndex,
			int fstCellColIndex, int secCellRowIndex, int secCellColIndex) {
		this.getTable(tableIndex);
		this.mergeCell(fstCellRowIndex, fstCellColIndex, secCellRowIndex,
				secCellColIndex);
	}

	public Dispatch getRangeParagraphs() throws Exception {
		return Dispatch.get(range, "Paragraphs").toDispatch();
	}

	public Dispatch getParagraph(int tIndex) throws Exception {
		return Dispatch.call(paragraphs, "Item", new Variant(tIndex))
				.toDispatch();
	}

	public Dispatch getParagraphRange() throws Exception {
		return Dispatch.get(paragraph, "range").toDispatch();
	}

	public String getRangeText() throws Exception {
		return Dispatch.get(range, "Text").toString();
	}

	public int getParagraphsCount() throws Exception {
		int count = 0;
		count = Dispatch.get(paragraphs, "Count").toInt();
		return count;
	}

	public static void main(String[] args) {
		long time1 = System.currentTimeMillis();
		int i = 0;
		ComThread.InitSTA(); // ��ʼ��com���̣߳��ǳ���Ҫ����ʹ�ý�����Ҫ���� realease����
		// Instantiate objWord //Declare word object
		ActiveXComponent objWord = new ActiveXComponent("Word.Application");
		// Assign a local word object
		Dispatch wordObject = (Dispatch) objWord.getObject();
		// Create a Dispatch Parameter to show the document that is opened
		Dispatch.put((Dispatch) wordObject, "Visible", new Variant(true)); // new
		// Variant(true)��ʾwordӦ�ó���ɼ�
		// Instantiate the Documents Property
		Dispatch documents = objWord.getProperty("Documents").toDispatch(); // documents��ʾword�������ĵ����ڣ���word�Ƕ��ĵ�Ӧ�ó���
		// Add a new word document, Current Active Document
		// Dispatch document = Dispatch.call(documents, "Add").toDispatch(); //
		// ʹ��Add�����һ�����ĵ�����Open������Դ�һ�������ĵ�
		
		Dispatch documentAll ;
		for (int n = 1; n <= 3; n++) {
			Dispatch document = Dispatch.call(documents, "Open", "d://file"+n+".doc")
					.toDispatch();
			Dispatch wordContent = Dispatch.get(document, "Content")
					.toDispatch(); // ȡ��word�ļ�������
			Dispatch.call(wordContent, "InsertAfter", ""); // ����һ������
			Dispatch.call(document, "SaveAs", new Variant("d://" + (i++)
					+ "_Final.doc")); // ����һ�����ĵ�
			Dispatch.call(document, "Close", new Variant(true));
		}
		//
		// end for
		Dispatch.call(objWord, "Quit");
		ComThread.Release(); // �ͷ�com�̡߳�����jacob�İ����ĵ���com���̻߳��ղ���java����������������
		long time2 = System.currentTimeMillis();
		double time3 = (time2 - time1) / 1000.0;
		System.out.println("/n" + time3 + " ��.");
	}
}
