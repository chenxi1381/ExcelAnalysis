package com.xchen.excelanalysis.bean;

public class DataBean {
	public String name = "null";
	public String value = "null";
	
	public String nameDif = "null";
	public String valueDif = "null";
	public String log;
	public DataBean(){}
	public DataBean(String name,String value,String nameDif,String valueDif){
		this.name = name;
		this.value = value;
		this.nameDif = nameDif;
		this.valueDif = valueDif;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	public String getNameDif() {
		return nameDif;
	}
	public void setNameDif(String nameDif) {
		this.nameDif = nameDif;
	}
	public String getValueDif() {
		return valueDif;
	}
	public void setValueDif(String valueDif) {
		this.valueDif = valueDif;
	}
	public String getLog() {
		return log;
	}
	public void setLog(String log) {
		this.log = log;
	}
	
	
}
