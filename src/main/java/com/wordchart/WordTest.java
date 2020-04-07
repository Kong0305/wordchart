package com.wordchart;

import java.io.IOException;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.policy.MyChartPolicy;

public class WordTest {
	
	public static void main(String[] args) throws IOException {
		String data = "{\r\n" + 
				"    \"khg\": [\r\n" + 
				"        [\r\n" + 
				"            \"这是第一个单元格\",\r\n" + 
				"            \"系列1\",\r\n" + 
				"            \"系列2\"\r\n" + 
				"        ],\r\n" + 
				"        [\r\n" + 
				"            \"一月\",\r\n" + 
				"            \"100\",\r\n" + 
				"            \"50\"\r\n" + 
				"        ],\r\n" + 
				"        [\r\n" + 
				"            \"二月\",\r\n" + 
				"            \"200\",\"100\"\r\n" + 
				"        ],\r\n" + 
				"        [\r\n" + 
				"            \"三月\",\r\n" + 
				"            \"300\",\"150\"\r\n" + 
				"        ],\r\n" + 
				"        [\r\n" + 
				"            \"四月\",\r\n" + 
				"            \"400\",\"200\"\r\n" + 
				"        ],\r\n" + 
				"        [\r\n" + 
				"            \"五月\",\r\n" + 
				"            \"500\",\"250\"\r\n" + 
				"        ],\r\n" + 
				"        [\r\n" + 
				"            \"六月\",\r\n" + 
				"            \"600\",\"300\"\r\n" + 
				"        ]\r\n" + 
				"    ]\r\n" + 
				"}";
		
		JSONObject params = JSON.parseObject(data);
		
		ConfigureBuilder configureBuilder = Configure.newBuilder();
		configureBuilder.bind("khg", new MyChartPolicy());
		
		Configure config = configureBuilder.build();
		
		
		// 核心API采用了极简设计，只需要一行代码
		XWPFTemplate.compile("D:\\2、工作目录\\2020年3月\\0323\\模板word-图表3.docx", config).render(params)
				.writeToFile("D:\\2、工作目录\\2020年3月\\0323\\生成word\\模板word-图表3-输出.docx");
	}

}
