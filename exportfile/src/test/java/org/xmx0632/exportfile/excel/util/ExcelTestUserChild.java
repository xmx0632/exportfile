package org.xmx0632.exportfile.excel.util;

import java.math.BigDecimal;
import java.util.Date;

import org.xmx0632.exportfile.excel.model.Excel;

public class ExcelTestUserChild extends ExcelTestUserParent{
	 
    @Excel(skip = true)
    private String password;
 
    @Excel(name = "xx")
    private Double xx;
 
    @Excel(name = "yy")
    private Date yy;
 
    @Excel(name = "锁定")
    private Boolean locked;
 
    @Excel(name = "金额")
    private BigDecimal db;

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public Double getXx() {
		return xx;
	}

	public void setXx(Double xx) {
		this.xx = xx;
	}

	public Date getYy() {
		return yy;
	}

	public void setYy(Date yy) {
		this.yy = yy;
	}

	public Boolean getLocked() {
		return locked;
	}

	public void setLocked(Boolean locked) {
		this.locked = locked;
	}

	public BigDecimal getDb() {
		return db;
	}

	public void setDb(BigDecimal db) {
		this.db = db;
	}
 
 
}