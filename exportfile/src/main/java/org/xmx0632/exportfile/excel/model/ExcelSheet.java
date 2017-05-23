package org.xmx0632.exportfile.excel.model;

import java.util.List;

public class ExcelSheet {

	private String title;
	private ExcelDataFormatter edf;
	private List content;

	public ExcelSheet() {
		super();
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public ExcelDataFormatter getEdf() {
		return edf;
	}

	public void setEdf(ExcelDataFormatter edf) {
		this.edf = edf;
	}

	public List getContent() {
		return content;
	}

	public void setContent(List content) {
		this.content = content;
	}
}
