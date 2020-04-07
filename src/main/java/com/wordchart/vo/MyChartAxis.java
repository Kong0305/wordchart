package com.wordchart.vo;

public class MyChartAxis {

	private String axisName;
	private String value;
	
	public MyChartAxis(String axisName,String value) {
		this.axisName = axisName;
		this.value = value;
	}

	public String getAxisName() {
		return axisName;
	}

	public void setAxisName(String axisName) {
		this.axisName = axisName;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

}
