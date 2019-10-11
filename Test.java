package com.zhiyou100.controller;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.zhiyou100.model.User;
import com.zhiyou100.service.UserService;

@Controller
public class POIController {

	
	@Autowired
	private UserService service;
	
	@Test
	@RequestMapping("/exportUser.do")
	public String exportUser(HttpServletResponse resp,
			@RequestParam Map<String, String> keywordMap)
					throws IOException {
		// 查全部用户
		List<User> users = service.findAllUser(keywordMap);
		System.out.println(users);
		// 导出
		//创建 工作表
		HSSFWorkbook wb = new HSSFWorkbook();
		// 创建sheet
		HSSFSheet sheet = wb.createSheet("用户信息");
		// 创建第一行
		HSSFRow r0 = sheet.createRow(0);
		// 创建第一列
		HSSFCell r0c0 = r0.createCell(0);
		// 设置内容
		r0c0.setCellValue("用户信息");
		// 设置单元格合并  
		/* 
		 * 参数1 : int firstRow, 从哪一行开始合并
		 * 参数2 : int lastRow, 到哪一行结束
		 * 参数3 : int firstCol, 从哪一列开始合并
		 * 参数4 : int lastCol, 到哪一列结束
		 * 
		 */
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
		
		// 创建第二行
		HSSFRow r1 = sheet.createRow(1);
		// 创建列,并设置值
		r1.createCell(0).setCellValue("用户编号");
		r1.createCell(1).setCellValue("用户名");
		r1.createCell(2).setCellValue("密码");
		r1.createCell(3).setCellValue("更新时间");
		r1.createCell(4).setCellValue("用户状态");
		r1.createCell(5).setCellValue("真实姓名");
		r1.createCell(6).setCellValue("邮箱");
		
		// 从第三行开始,就是从数据库查出的数据
		for(int i = 0;i<users.size();i++) {
			HSSFRow row = sheet.createRow(i+2); // 从第三行开始创建行
			// 创建列.并赋值
			row.createCell(0).setCellValue(users.get(i).getId());
			row.createCell(1).setCellValue(users.get(i).getUser_name());
			row.createCell(2).setCellValue(users.get(i).getPassword());
			row.createCell(3).setCellValue(users.get(i).getUpdate_time());
			row.createCell(4).setCellValue(users.get(i).getStatus());
			row.createCell(5).setCellValue(users.get(i).getReal_name());
			row.createCell(6).setCellValue(users.get(i).getEmail());
		}
		// 解决响应中文文件名乱码问题
		String filename = URLEncoder.encode("用户信息表", "utf-8");
		// 浏览器响应下载弹框
		resp.setHeader("Content-disposition", "attachment;filename="+filename+".xls");
		resp.setContentType("application/msexcel");
		// 输出
		OutputStream out = resp.getOutputStream();
		wb.write(out);
		out.close();
		
		return null;
	}

}
