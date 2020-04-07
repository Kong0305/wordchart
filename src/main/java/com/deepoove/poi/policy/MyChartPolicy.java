package com.deepoove.poi.policy;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;

import com.alibaba.fastjson.JSONArray;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.MyRunTemplate;
import com.wordchart.vo.MyChartAxis;
import com.wordchart.vo.MyChartSeries;

public class MyChartPolicy extends AbstractRenderPolicy<JSONArray> {

	@Override
	public void doRender(RenderContext<JSONArray> context) throws Exception {
		ElementTemplate elementTemplate = context.getEleTemplate();
		if (elementTemplate != null && elementTemplate instanceof MyRunTemplate) {
			MyRunTemplate myRunTemplate = (MyRunTemplate) elementTemplate;
			XWPFChart chart = myRunTemplate.getChart();
			if (chart != null) {
				List<MyChartSeries> serList = this.resolveDatas(context.getData());
				String sheetName = this.refreshExcel(chart, context.getData());
				this.refreshGraphContent(chart, sheetName, context.getData(), serList);
			}
		}
	}

	/**
	 * 刷新EXCEL数据
	 * 
	 * @param chart
	 * @param rows
	 * @return sheet名
	 */
	private String refreshExcel(XWPFChart chart, JSONArray rows) {
		String sheetName = null;
		if (chart == null || rows == null) {
			return sheetName;
		}
		try {
			XSSFWorkbook workbook = chart.getWorkbook();
			// 获取原sheet名
			sheetName = workbook.getSheetName(0);
			// 删除原有sheet
			workbook.removeSheetAt(0);
			// 根据原有sheet名新创建sheet
			Sheet sheet = workbook.createSheet(sheetName);
			this.createRows(rows, sheet);
			return sheetName;
		} catch (Exception e) {
			e.printStackTrace();
			return sheetName;
		}
	}

	/**
	 * 创建行数据
	 * 
	 * @param rows
	 * @param sheet
	 */
	private void createRows(JSONArray rows, Sheet sheet) {
		if (rows == null || sheet == null) {
			return;
		}
		for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
			JSONArray cols = rows.getJSONArray(rowIndex);
			Row row = sheet.createRow(rowIndex);
			this.createCells(cols, row);
		}
	}

	/**
	 * 创建单元格数据
	 * 
	 * @param cols
	 * @param row
	 */
	private void createCells(JSONArray cols, Row row) {
		if (cols == null || row == null) {
			return;
		}
		int rowNum = row.getRowNum();
		for (int colIndex = 0; colIndex < cols.size(); colIndex++) {
			String cellValueStr = cols.getString(colIndex);
			// 首行、首列分别代表系列名、横坐标,非首行首列均是数值
			// 避免值中可能存在非数值,因此做下数值判断
			if (rowNum == 0 || colIndex == 0 || !NumberUtils.isNumber(cellValueStr)) {
				row.createCell(colIndex).setCellValue(cellValueStr);
			} else {
				Double cellValue = cols.getDouble(colIndex);
				row.createCell(colIndex).setCellValue(cellValue);
			}

		}
	}

	/**
	 * 将数据封装成MyChartSeries对象,便于word中图表解析时使用 数据格式参考<br>
	 * (留空) 系列1 系列2<br>
	 * 第一季度 5 10<br>
	 * 第二季度 10 15<br>
	 * 第三季度 15 16<br>
	 * 第四季度 20 4<br>
	 * 
	 * @param rows
	 */
	private List<MyChartSeries> resolveDatas(JSONArray rows) {
		List<MyChartSeries> serAxisList = new ArrayList<MyChartSeries>();
		// 数据为空
		if (CollectionUtils.isEmpty(rows)) {
			return serAxisList;
		}

		// 第0行没有数据
		JSONArray serJsonArray = rows.getJSONArray(0);
		if (CollectionUtils.isEmpty(serJsonArray)) {
			return serAxisList;
		}

		for (int serIndex = 1; serIndex < serJsonArray.size(); serIndex++) {
			String seriesName = serJsonArray.getString(serIndex);
			MyChartSeries myChartData = new MyChartSeries(seriesName);
			serAxisList.add(myChartData);
		}

		// 第0行数据不全,没有系列名
		if (CollectionUtils.isEmpty(serAxisList)) {
			return serAxisList;
		}

		// 第一行开始,第0个单元格代表axis(横坐标)名称,后续的依次为每个系列数值
		for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
			JSONArray cols = rows.getJSONArray(rowIndex);
			String axisName = null;
			if (cols != null && cols.size() > 0) {
				axisName = cols.getString(0);
			}

			for (int serIndex = 0; serIndex < serAxisList.size(); serIndex++) {
				MyChartSeries myChartData = serAxisList.get(serIndex);

				int colIndex = serIndex + 1;
				String colValue = null;
				if (cols != null && cols.size() > colIndex) {
					colValue = cols.getString(colIndex);
				}
				MyChartAxis myChartAxis = new MyChartAxis(axisName, colValue);
				myChartData.getAxisDataList().add(myChartAxis);
			}
		}
		return serAxisList;
	}

	/**
	 * 刷新图表数据
	 * 
	 * @param chart
	 * @param sheetName
	 * @param rows
	 * @param serList
	 */
	private void refreshGraphContent(XWPFChart chart, String sheetName, JSONArray rows, List<MyChartSeries> serList) {
		CTChart ctChart = chart.getCTChart();
		CTBarChart ctBarChart = ctChart.getPlotArea().getBarChartArray(0);

		// 原有的所有系列
		List<CTBarSer> ctBarSers = ctBarChart.getSerList();
		// 按照新系列的数量,对应删减老系列的数量;保留老系列的原因是为了尽可能的保留原有系列样式
		if (ctBarSers != null) {
			List<CTBarSer> newCtBarSers = ctBarSers.subList(0, Math.min(serList.size(), ctBarSers.size()));
			CTBarSer[] newCtBarSersArray = new CTBarSer[newCtBarSers.size()];
			newCtBarSers.toArray(newCtBarSersArray);
			ctBarChart.setSerArray(newCtBarSersArray);
		}

		// 构造每个系列的序列、数据、系列名
		for (int serIndex = 0; serIndex < serList.size(); serIndex++) {
			MyChartSeries myChartData = serList.get(serIndex);
			CTBarSer ctBarSer = null;
			if (ctBarSers != null && ctBarSers.size() > serIndex) {
				ctBarSer = ctBarChart.getSerArray(serIndex);
			}
			if (ctBarSer == null) {
				ctBarSer = ctBarChart.addNewSer();
			} else {
				ctBarSer.unsetCat();
				ctBarSer.unsetVal();
				ctBarSer.unsetTx();
			}

			// Category Axis Data
			CTAxDataSource cat = ctBarSer.addNewCat();
			// 获取图表的值
			CTNumDataSource val = ctBarSer.addNewVal();
			// 系列名称
			CTSerTx ctSerTx = ctBarSer.addNewTx();

			CTStrData strData = cat.addNewStrRef().addNewStrCache();
			CTNumData numData = val.addNewNumRef().addNewNumCache();
			CTStrData txData = ctSerTx.addNewStrRef().addNewStrCache();

			// 构造序列项、数据
			int idx = 0;
			for (MyChartAxis mChartAxis : myChartData.getAxisDataList()) {
				String axisName = mChartAxis.getAxisName();
				String value = mChartAxis.getValue();

				CTStrVal sVal = strData.addNewPt();// 序列名称
				sVal.setIdx(idx);
				sVal.setV(axisName);

				CTNumVal numVal = numData.addNewPt();// 序列值
				numVal.setIdx(idx);
				numVal.setV(value);
				++idx;
			}
			// 设置系列名称
			CTStrVal txVal = txData.addNewPt();
			txVal.setIdx(0);
			txVal.setV(myChartData.getSeries());

			numData.addNewPtCount().setVal(idx);
			strData.addNewPtCount().setVal(idx);

			// 序列区域
			String axisDataRange = new CellRangeAddress(1, rows.size() - 1, 0, 0).formatAsString(sheetName, true);
			cat.getStrRef().setF(axisDataRange);

			// 数据区域
			String numDataRange = new CellRangeAddress(1, rows.size() - 1, serIndex + 1, serIndex + 1)
					.formatAsString(sheetName, true);
			val.getNumRef().setF(numDataRange);

			// 系列名区域
			String serDataRange = new CellRangeAddress(0, 0, serIndex + 1, serIndex + 1).formatAsString(sheetName,
					true);
			ctSerTx.getStrRef().setF(serDataRange);
		}

	}

}
