package com.jun.controller;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelController extends HttpServlet {
	/**
	 * 将待转换Excel此路径
	 */
	public String sourcePath = "E:/ChengeExcel/";

	/**
	 * 从第几行开始读取数据(第一行为0)
	 */
	public int startRowNum = 1;
	
	/**
	 * 文件名
	 */
	public String fileName = "";
	
	/**
	 * action
	 */
	public String action = "";

	/**
	 * 处理线程
	 */
	public Thread run = null;

	/**
	 * 最后一行和最后一列
	 */
	public int lastRowNum, lastCellNum = 0;

	/**
	 * 运行是否异常
	 */
	public boolean flag = false;
	
	/**
	 * 是否运行完成
	 */
	public boolean state = false;

	/**
	 * 异常信息
	 */
	public String msg = "";
	
	/**
	 * 核心解析逻辑
	 */
	public OpeningByXls open;
	
	/**
	 * 核心解析逻辑2
	 */
	public OpeningJarDataByXls open2;

	@Override
	protected void doPost(HttpServletRequest req, HttpServletResponse resp)
			throws ServletException, UnsupportedEncodingException {
		if (run != null && run.isAlive()) {
			flag = true;
			msg = "后台正在处理：" + fileName + "中，请稍后...";
			return;
		} else {
			init();
		}
		req.setCharacterEncoding("utf-8");
		fileName = req.getParameter("fileName");
		action = req.getParameter("action");
		Runnable parse = new Runnable() {
			public void run() {
				try {
					List<Map<String, String>> data = parseExcel(sourcePath + fileName);
					if("A".equals(action)){
						open = new OpeningByXls();
						open.initIquantityDate(data);
					}else if("B".equals(action)){
						open2 = new OpeningJarDataByXls();
						open2.initIquantityDate(data);
					}
					state = true;
				} catch (Exception e) {
					flag = true;
					msg = e.getMessage();
					e.printStackTrace();
				}
			}
		};
		run = new Thread(parse);
		run.start();
	}

	@Override
	protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
		resp.setCharacterEncoding("utf-8");
		PrintWriter out = resp.getWriter();
		if (flag) {
			out.print("Msg：" + msg);
		}else if(state){
			if(open != null){
				out.print("解析完成!用时：" + open.getTime());
			}else if(open2 != null){
				out.print("解析完成!用时：" + open2.getTime());
			}
		}else{
			if(open != null){
				out.print("当前进度：" + open.getProgress());
			}else if(open2 != null){
				out.print("当前进度：" + open2.getProgress());
			}
		}
		out.close();
	}

	public List<Map<String, String>> parseExcel(String filePath) throws Exception {
		File file = new File(filePath);
		// 创建Excel对象，读取文件
		Workbook workbook = null;
		if (filePath.endsWith("xls")) {
			workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
		} else if (filePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook(FileUtils.openInputStream(file));
		} else {
			throw new Exception("文档类型错误！");
		}
		// 通过名字“Sheet0”获取工作表
		// HSSFSheet sheet = workbook.getSheet("Sheet0");
		// 读取默认第一个工作表sheet
		Sheet sheet = workbook.getSheetAt(0);
		// 最后一行行号
		lastRowNum = sheet.getLastRowNum();
		//将excle转换成List
		List<Map<String, String>> excelData = new ArrayList<Map<String, String>>();
		// 读取每一行
		for (int i = startRowNum; i < lastRowNum + 1; i++) {
			Row row = sheet.getRow(i);
			// 获取当前行最后单元格列号
			lastCellNum = row.getLastCellNum();
			// 每一列存进Map
			Map<String, String> rowData = new HashMap<String ,String>();
			// 读取该行每一个cell
			for (int j = 0; j < lastCellNum; j++) {
				rowData.put(getColumnCharName(j), parseRow(row, j));
			}
			rowData.put("RN", i + 1 + "");
			excelData.add(rowData);
		}
		workbook.close();
		return excelData;
	}
	
	public String parseRow(Row row, int j) throws Exception {
		Cell cell = row.getCell(j);
		if (cell != null) {
			return cell.getStringCellValue();
		}
		return null;
	}

	public String getColumnCharName(int index){
		return String.valueOf((char) (65 + index));
	}
	
	public void init() {
		fileName = "";
		lastRowNum = 0;
		lastCellNum = 0;
		flag = false;
		msg = "";
		state = false;
	}

}
