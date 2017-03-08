package com.whl.poi.demo;

import java.math.BigDecimal;

import com.whl.poi.Column;
import com.whl.poi.RowData;

@RowData(rowAcross = 2, desc = "用户数据")
public class UserInfo {

	@Column(colIndex = 0, isAcross = true, format = "0")
	private int id;
	@Column(colIndex = 1, isAcross = true)
	private String name;

	@Column(colIndex = 2, isAcross = true)
	private String cardNo;

	@Column(colIndex = 3, isAcross = true, format = "#,##0.00")
	private BigDecimal salary;

	@Column(colIndex = 4, rowIndex = 0)
	private String personInsTitle;
	
	@Column(colIndex = 4, rowIndex = 1)
	private String firmInsTitle;

	@Column(colIndex = 5, rowIndex = 0, format = "#,##0.00")
	private BigDecimal personIns;

	@Column(colIndex = 5, rowIndex = 1, format = "#,##0.00")
	private BigDecimal firmIns;

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getCardNo() {
		return cardNo;
	}

	public void setCardNo(String cardNo) {
		this.cardNo = cardNo;
	}

	public BigDecimal getSalary() {
		return salary;
	}

	public void setSalary(BigDecimal salary) {
		this.salary = salary;
	}

	public String getPersonInsTitle() {
		return personInsTitle;
	}

	public void setPersonInsTitle(String personInsTitle) {
		this.personInsTitle = personInsTitle;
	}

	public String getFirmInsTitle() {
		return firmInsTitle;
	}

	public void setFirmInsTitle(String firmInsTitle) {
		this.firmInsTitle = firmInsTitle;
	}

	public BigDecimal getPersonIns() {
		return personIns;
	}

	public void setPersonIns(BigDecimal personIns) {
		this.personIns = personIns;
	}

	public BigDecimal getFirmIns() {
		return firmIns;
	}

	public void setFirmIns(BigDecimal firmIns) {
		this.firmIns = firmIns;
	}

	public static void main(String[] args) throws Exception {
		RowData a = UserInfo.class.getAnnotation(RowData.class);
		System.out.println(a.desc());
	}

}
