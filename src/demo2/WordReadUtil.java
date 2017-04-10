package demo2;
/**
 * @author xww
 *
 */
import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;

public class WordReadUtil {

	private static  String path = "E:\\testExcel\\";

	/**
	 * 此标记用于取到我们需要的模块内容
	 */
	private static int flag = 0;
	
	
	public static String getPath(String filename) {
		return path+filename;
	}
	public static void main(String[] args) throws FileNotFoundException,
			IOException, XmlException, OpenXML4JException {
		flagWordVersion(path, "2", "代码规范");
	}

	/**
	 * 判断word是哪个版本的
	 * 
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws XmlException
	 */
	public static List<String> flagWordVersion(String path, String n, String obj)
			throws IOException, XmlException, OpenXML4JException {
		List<String> list =null;
		InputStream inp = new FileInputStream(path);
		if (!inp.markSupported()) {
			inp = new PushbackInputStream(inp, 8);
		}
		if (POIFSFileSystem.hasPOIFSHeader(inp)) {
			if (n == "" || n == null)
				list=getWordTitles2003(path);
			else
				list=getWordTitles2003(path, n, obj);
		}
		if (POIXMLDocument.hasOOXMLHeader(inp)) {
			if (n == "" || n == null)
				list=getWordTitles2007(path);
			else
				list=getWordTitles2007(path, n, obj);
		}
		return list;
	}

	/**
	 * @param path
	 * @return
	 * @throws IOException
	 * @desc 显示全部标题
	 */
	public static List<String> getWordTitles2003(String path)
			throws IOException {

		File file = new File(path);

		String filename = file.getName();

		filename = filename.substring(0, filename.lastIndexOf("."));

		InputStream is = new FileInputStream(path);

		HWPFDocument doc = new HWPFDocument(is);

		Range r = doc.getRange();
		List<String> list = new ArrayList<String>();
		// 用来获取段落数
		for (int i = 0; i < r.numParagraphs(); i++) {
			// 获取目标索引段落
			Paragraph p = r.getParagraph(i);
			// check if style index is greater than total number of styles

			int numStyles = doc.getStyleSheet().numStyles();
			int styleIndex = p.getStyleIndex();

			if (numStyles > styleIndex) {

				StyleSheet style_sheet = doc.getStyleSheet();

				StyleDescription style = style_sheet
						.getStyleDescription(styleIndex);
				String styleName = style.getName();
				if (styleName != null && styleName.equals("标题 ")) {
					String text = p.text();
					list.add(text);
				}
			}
		}
		// 得到word数据流
		byte[] dataStream = doc.getDataStream();
		// 用于在一段范围内获得段落数
		int numCharacterRuns = r.numCharacterRuns();
		// System.out.println("CharacterRuns 数:"+numCharacterRuns);
		/*// 负责图像提取 和 确定一些文件某块是否包含嵌入的图像。
		System.out.println(list + "--------");*/
		return list;
	}

	public static List<String> getWordTitles2007(String path)
			throws IOException, XmlException, OpenXML4JException {
		InputStream is = new FileInputStream(path);
		OPCPackage p = POIXMLDocument.openPackage(path);
		XWPFWordExtractor e = new XWPFWordExtractor(p);
		POIXMLDocument doc = e.getDocument();
		List<String> list = new ArrayList<String>();
		XWPFDocument doc1 = new XWPFDocument(is);
		List<XWPFParagraph> paras = doc1.getParagraphs();
		for (XWPFParagraph graph : paras) {
			String text = graph.getParagraphText();
			String style = graph.getStyle();
			list.add(text);
		}
		return list;
	}

	/**
	 * 代表N级目录
	 * 
	 * @param path
	 * @param n
	 * @param obj
	 * @return
	 * @throws IOException
	 * @desc obj 代表按个标题,意思就是N级目录的上一级标题头为obj的模块
	 */
	public static List<String> getWordTitles2003(String path, String n,
			String obj) throws IOException {

		File file = new File(path);

		String filename = file.getName();

		filename = filename.substring(0, filename.lastIndexOf("."));

		InputStream is = new FileInputStream(path);

		HWPFDocument doc = new HWPFDocument(is);

		Range r = doc.getRange();
		List<String> list = new ArrayList<String>();
		// 用来获取段落数
		for (int i = 0; i < r.numParagraphs(); i++) {
			// 获取目标索引段落
			Paragraph p = r.getParagraph(i);
			// check if style index is greater than total number of styles

			int numStyles = doc.getStyleSheet().numStyles();

			int styleIndex = p.getStyleIndex();

			if (numStyles > styleIndex) {

				StyleSheet style_sheet = doc.getStyleSheet();
				StyleDescription style = style_sheet
						.getStyleDescription(styleIndex);
				String styleName = style.getName();
				// 前后2个if用以判断是那个节点
				if (obj == null || "".equals(obj)) {
					if (styleName != null && styleName.equals("标题 " + n)) {
						String text = p.text();
						list.add(text);
					}
				} else {
					aopRead2003(styleName, n, p, obj, list);
				}
			}
		}
		// 得到word数据流
		byte[] dataStream = doc.getDataStream();
		// 用于在一段范围内获得段落数
		int numCharacterRuns = r.numCharacterRuns();
		// System.out.println("CharacterRuns 数:"+numCharacterRuns);
		// 负责图像提取 和 确定一些文件某块是否包含嵌入的图像。
		System.out.println(list + "--------");
		return list;
	}

	/**
	 * 封装一部分代码，不然看的 心烦
	 * 
	 * @param styleName
	 * @param n
	 * @param p
	 * @param obj
	 * @param list
	 */
	public static void aopRead2003(String styleName, String n, Paragraph p,
			String obj, List<String> list) {
		if (styleName!=null &&styleName.equals("标题 " + String.valueOf(Integer.parseInt(n) - 1))) {
			if (p.text().contains(obj))
				flag++;
		}
		if (styleName != null && styleName.equals("标题 " + n) && flag > 0) {
			String text = p.text();
			list.add(text);
		}
		if (styleName != null && styleName.equals("标题 " + String.valueOf(Integer.parseInt(n) - 1))) {
			if (!p.text().contains(obj))
				flag = 0;
		}
	}

	/**
	 * n代表N级目录
	 * 
	 * @param path
	 * @param n
	 * @return
	 * @throws IOException
	 * @throws XmlException
	 * @throws OpenXML4JException
	 * @desc 目前支持3级目录
	 */
	public static List<String> getWordTitles2007(String path, String n,
			String obj) throws IOException, XmlException, OpenXML4JException {
		n=String.valueOf(Integer.parseInt(n)+1);
		InputStream is = new FileInputStream(path);
		OPCPackage p = POIXMLDocument.openPackage(path);
		XWPFWordExtractor e = new XWPFWordExtractor(p);
		POIXMLDocument doc = e.getDocument();
		List<String> list = new ArrayList<String>();
		XWPFDocument doc1 = new XWPFDocument(is);
		List<XWPFParagraph> paras = doc1.getParagraphs();
		if ("".equals(obj) || null == obj) {
			aopRead2007(paras, n, obj, list);
		} else {
			aopRead20072(paras, n, obj, list);
		}
		return list;
	}

	/**
	 * 
	 * @param paras
	 * @param n
	 * @param p
	 * @param obj
	 * @param list
	 */
	public static void aopRead2007(List<XWPFParagraph> paras, String n,
			String obj, List<String> list) {
		for (XWPFParagraph graph : paras) {
			String text = graph.getParagraphText();
			String style = graph.getStyle();
			if (n.equals(style)) {
				list.add(text);
			} else {
				continue;
			}

		}
	}

	/**
	 * 
	 * @param paras
	 * @param n
	 * @param p
	 * @param obj
	 * @param list
	 */
	public static void aopRead20072(List<XWPFParagraph> paras, String n,
			String obj, List<String> list) {
		for (XWPFParagraph graph : paras) {
			String text = graph.getParagraphText();
			String style = graph.getStyle();
			if (style!=null&&style.equals(String.valueOf(Integer.parseInt(n) - 1))) {
				if (text.contains(obj))
					flag++;
			}
			if (style!=null&&n.equals(style) && flag > 0 ) {
				System.out.println(text);
				list.add(text);
			}
			if (style!=null&&style.equals(String.valueOf(Integer.parseInt(n) - 1))) {
				if (!text.contains(obj))
					flag = 0;
			}
		}
		System.out.println(list);
	}
}
