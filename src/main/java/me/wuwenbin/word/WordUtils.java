package me.wuwenbin.word;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Bookmark;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 
 * <b>ClassName</b>: WordUtils<br>
 * <b>Description</b>: 有关word的读取的工具类，大致2个分类：07版本以下（含），07版本以上（不含）<br>
 * 本工具类自是一些基础的word操作，更详细的请自行使用apache的poi相关类功能。去实现<br>
 * <b>Version</b>: Ver 1.0<br>
 *
 * <b>author</b>: Wuwenbin<br>
 * <b>date</b>: 2016年8月30日<br>
 * <b>time</b>: 下午4:25:17 <br>
 */
public class WordUtils {

	/**
	 * 获取WordExtrator对象，私有方法避免重复编写
	 * 
	 * @param path
	 * @return
	 * @throws IOException
	 */
	private static WordExtractor getWE(String path) throws IOException {
		InputStream inputStream = new FileInputStream(path);
		return new WordExtractor(inputStream);
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocWholeTxt<br>
	 * <b>Description</b> : 获取doc文档的全部文本信息<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	public static String getDocWholeTxt(String docPath) throws IOException {
		return getWE(docPath).getText();
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocParagraphText<br>
	 * <b>Description</b> : 获取doc文档的每个段落的信息<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	public static String[] getDocParagraphText(String docPath) throws IOException {
		return getWE(docPath).getParagraphText();
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocHeaderTxt<br>
	 * <b>Description</b> : 获取doc文档的页眉信息<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	@SuppressWarnings("deprecation")
	public static String getDocHeaderTxt(String docPath) throws IOException {
		return getWE(docPath).getHeaderText();
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocFooterTxt<br>
	 * <b>Description</b> : 获取doc文档的页脚信息<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	@SuppressWarnings("deprecation")
	public static String getDocFooterTxt(String docPath) throws IOException {
		return getWE(docPath).getFooterText();
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocMetaInfo<br>
	 * <b>Description</b> : 获取当前word文档的元数据信息，包括作者、文档的修改时间等<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	public static String getDocMetaInfo(String docPath) throws IOException {
		return getWE(docPath).getMetadataTextExtractor().getText();
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocInfoMeta<br>
	 * <b>Description</b> : 获取word文档中的一些基础元信息<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	public WordInfoMeta getDocInfoMeta(String docPath) throws IOException {
		SummaryInformation si = getWE(docPath).getSummaryInformation();
		DocumentSummaryInformation dsi = getWE(docPath).getDocSummaryInformation();
		WordInfoMeta bwim = new WordInfoMeta();
		bwim.setAuthor(si.getAuthor()).setCharCount(bwim.getCharCount()).setPageCount(si.getPageCount()).setTitle(si.getTitle()).setSubject(si.getSubject());
		bwim.setCategory(dsi.getCategory()).setCompany(dsi.getCompany());
		return bwim;
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocBookmarks<br>
	 * <b>Description</b> : 获取word文档书签<br>
	 * 
	 * @param docPath
	 * @return
	 * @throws IOException
	 */
	public static List<Bookmark> getDocBookmarks(String docPath) throws IOException {
		InputStream is = new FileInputStream(docPath);
		HWPFDocument doc = new HWPFDocument(is);
		int count = doc.getBookmarks().getBookmarksCount();
		List<Bookmark> bookmarks = new ArrayList<Bookmark>();
		for (int i = 0; i < count; i++) {
			Bookmark bookmark = doc.getBookmarks().getBookmark(i);
			bookmarks.add(bookmark);
		}
		return bookmarks;
	}

	// ==================以下为读取docx文档======================

	public static enum CorePropType {
		/**
		 * 分类
		 */
		CATE,
		/**
		 * 创建者
		 */
		CREATER,
		/**
		 * 创建时间
		 */
		CTEATED,
		/**
		 * 标题
		 */
		TITLE,
		/**
		 * 主题
		 */
		SUBJECT;
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocxWholeTxt<br>
	 * <b>Description</b> : 读取docx文档文本信息<br>
	 * 
	 * @param docxPath
	 * @return
	 * @throws IOException
	 */
	public static String getDocxWholeTxt(String docxPath) throws IOException {
		InputStream is = new FileInputStream(docxPath);
		XWPFDocument doc = new XWPFDocument(is);
		XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
		return extractor.getText();
	}

	/**
	 * 
	 * <b>Author</b> : Wuwenbin<br>
	 * <b>Title</b> : getDocxCorePropertiesByType<br>
	 * <b>Description</b> : 获取一些主要信息：如分类、创建者、创建时间、主题、标题等。<br>
	 * 
	 * @param docxPath
	 * @param corePropType
	 * @return
	 * @throws IOException
	 */
	public static String getDocxCorePropertiesByType(String docxPath, CorePropType corePropType) throws IOException {
		InputStream is = new FileInputStream(docxPath);
		XWPFDocument doc = new XWPFDocument(is);
		XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
		CoreProperties coreProps = extractor.getCoreProperties();
		switch (corePropType) {
		case CATE:
			return coreProps.getCategory();
		case CREATER:
			return coreProps.getCreator();
		case CTEATED:
			return String.valueOf(coreProps.getCreated().getTime());
		case TITLE:
			return coreProps.getTitle();
		case SUBJECT:
			return coreProps.getSubject();
		default:
			return "";
		}
	}
}
