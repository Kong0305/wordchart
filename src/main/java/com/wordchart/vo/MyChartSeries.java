package com.wordchart.vo;

import java.util.ArrayList;
import java.util.List;

public class MyChartSeries {
	// 系列名称
	private String series;

	private List<MyChartAxis> axisDataList = new ArrayList<MyChartAxis>();

	public MyChartSeries(String series) {
		this.series = series;
	}

	public String getSeries() {
		return series;
	}

	public List<MyChartAxis> getAxisDataList() {
		return axisDataList;
	}
}
