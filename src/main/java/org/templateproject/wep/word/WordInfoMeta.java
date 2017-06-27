package org.templateproject.wep.word;

/**
 * 
 * <b>ClassName</b>: WordInfoMeta<br>
 * <b>Description</b>: Word文档的一些基本的元信息<br>
 * <b>Version</b>: Ver 1.0<br>
 *
 * <b>author</b>: Wuwenbin<br>
 * <b>date</b>: 2016年8月30日<br>
 * <b>time</b>: 下午8:50:49 <br>
 */
public class WordInfoMeta {

	private String author;// 作者
	private String charCount;// 字符统计
	private int pageCount;// 页数
	private String title; // 标题
	private String subject;// 主题
	private String category; // 分类
	private String company; // 公司

	public String getAuthor() {
		return author;
	}

	public WordInfoMeta setAuthor(String author) {
		this.author = author;
		return this;
	}

	public String getCharCount() {
		return charCount;
	}

	public WordInfoMeta setCharCount(String charCount) {
		this.charCount = charCount;
		return this;
	}

	public int getPageCount() {
		return pageCount;
	}

	public WordInfoMeta setPageCount(int pageCount) {
		this.pageCount = pageCount;
		return this;
	}

	public String getTitle() {
		return title;
	}

	public WordInfoMeta setTitle(String title) {
		this.title = title;
		return this;
	}

	public String getSubject() {
		return subject;
	}

	public WordInfoMeta setSubject(String subject) {
		this.subject = subject;
		return this;
	}

	public String getCategory() {
		return category;
	}

	public WordInfoMeta setCategory(String category) {
		this.category = category;
		return this;
	}

	public String getCompany() {
		return company;
	}

	public WordInfoMeta setCompany(String company) {
		this.company = company;
		return this;
	}
}
