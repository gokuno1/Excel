package com.example.Excel.Reports.Models;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "north_login_details")
public class NReport {
	
	@Column(name="ZONE")
	private String zone;
	
	@Column(name="REGION")
	private String region;
	
	@Column(name="DEPOT")
	private String depot;
	
	@Column(name="ROLE")
	private String role;
	
	@Column(name="SAP_CODE")
	private String sapcode;
	
	@Id
	@Column(name="MOBILE_NUMBER")
	private String mobileNo;
	
	@Column(name="COMPANY_NAME")
	private String companyName;
	
	@Column(name="FIRST_NAME")
	private String firstName;
	
	@Column(name="LAST_NAME")
	private String lastName;
	
	@Column(name="LAST_LOGIN_DATE")
	private String lastLoginDate;
	
	@Column(name="FIRST_TIME_LOGIN_DATE")
	private String firstTimeLoginDate;

	public String getZone() {
		return zone;
	}

	public void setZone(String zone) {
		this.zone = zone;
	}

	public String getRegion() {
		return region;
	}

	public void setRegion(String region) {
		this.region = region;
	}

	public String getDepot() {
		return depot;
	}

	public void setDepot(String depot) {
		this.depot = depot;
	}

	public String getRole() {
		return role;
	}

	public void setRole(String role) {
		this.role = role;
	}

	public String getSapcode() {
		return sapcode;
	}

	public void setSapcode(String sapcode) {
		this.sapcode = sapcode;
	}

	public String getMobileNo() {
		return mobileNo;
	}

	public void setMobileNo(String mobileNo) {
		this.mobileNo = mobileNo;
	}

	public String getCompanyName() {
		return companyName;
	}

	public void setCompanyName(String companyName) {
		this.companyName = companyName;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public String getLastLoginDate() {
		return lastLoginDate;
	}

	public void setLastLoginDate(String lastLoginDate) {
		this.lastLoginDate = lastLoginDate;
	}

	public String getFirstTimeLoginDate() {
		return firstTimeLoginDate;
	}

	public void setFirstTimeLoginDate(String firstTimeLoginDate) {
		this.firstTimeLoginDate = firstTimeLoginDate;
	}

	
}
