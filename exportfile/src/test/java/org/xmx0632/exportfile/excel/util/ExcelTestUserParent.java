package org.xmx0632.exportfile.excel.util;

import org.xmx0632.exportfile.excel.model.Excel;

public class ExcelTestUserParent {
	 
    @Excel(name = "姓名", width = 30)
    protected String name;
 
    @Excel(name = "年龄", width = 60)
    protected String age;
 

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAge() {
		return age;
	}

	public void setAge(String age) {
		this.age = age;
	}

}