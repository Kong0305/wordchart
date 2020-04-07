package com.deepoove.poi.template.run;

import org.apache.poi.xwpf.usermodel.XWPFChart;

public class MyRunTemplate extends RunTemplate {
	private XWPFChart chart;

	public XWPFChart getChart() {
		return chart;
	}

	public void setChart(XWPFChart chart) {
		this.chart = chart;
	}

}
