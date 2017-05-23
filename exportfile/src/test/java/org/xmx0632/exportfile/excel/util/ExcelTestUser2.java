package org.xmx0632.exportfile.excel.util;

import org.xmx0632.exportfile.excel.model.Excel;

public class ExcelTestUser2 {
	 
    @Excel(name = "姓名", width = 30)
    private String name;
 
    @Excel(name = "年龄", width = 60)
    private String age;
 
    @Excel(skip = true)
    private String password;
 

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

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

}